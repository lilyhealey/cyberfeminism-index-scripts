// CONSTANTS
const ATTRIBUTES = [
  'TITLE',
  'SUBTITLE',
  'BODY',
  'ANNOTATION',
  'CROSSREF',
  'BOLD',
  'ITALIC'
];

const MAP = {
  TITLE:      { start: '<title>',      end: '</title>'},
  SUBTITLE:   { start: '<subtitle>',   end: '</subtitle>'},
  BODY:       { start: '<body>',   end: '</body>'},
  ANNOTATION: { start: '<annotation>', end: '</annotation>'},
  CROSSREF:   { start: '<crossref>',   end: '</crossref>'},
  BOLD:       { start: '<bold>',       end: '</bold>'    },
  ITALIC:     { start: '<italic>',     end: '</italic>'  },
};

const FORMATTING = [
  {
    name: 'TITLE',
    style: {
      fgColours: null,
      bgColours: [
        '#f4cccc',
        '#e6b8af'
      ],
    }
  },
  {
    name: 'SUBTITLE',
    style: {
      fgColours: null,
      bgColours: [
        '#d9ead3'
      ],
    }
  },
  {
    name: 'ANNOTATION',
    style: {
      fgColours: null,
      bgColours: [
        '#c9daf8',
        '#cfe2f3'
      ],
    }
  },
  {
    name: 'CROSSREF',
    style: {
      fgColours: [
        '#cccccc',
        '#b7b7b7'                
      ],
      bgColours: null,
    }
  },
  {
    name: 'BODY',
    style: {
      fgColours: null,
      bgColours: [ null, '#ffffff' ],
    }
  }
];

/** 
 * Adds the menu item on open. 
 */
function onOpen(e) {
  const ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('Convert to XML', 'starterXML')
    .addItem('Convert to TXT', 'starterTXT')
    .addToUi();
}

/** 
 * Returns an HTML dialog with the download link.
 */
function starterXML() {
  const template = HtmlService.createTemplateFromFile('dialog');
  template.syntax = 'xml';
  const html = template.evaluate();
  html.setWidth(300)
    .setHeight(150); 
  DocumentApp.getUi().showModalDialog(html, 'XML Conversion');
}

/** 
 * Returns an HTML dialog with the download link.
 */
function starterTXT() {
  const template = HtmlService.createTemplateFromFile('dialog');
  template.syntax = 'txt';
  const html = template.evaluate();
  html.setWidth(300)
    .setHeight(150); 
  DocumentApp.getUi().showModalDialog(html, 'TXT Conversion');
}

/** 
 * This function is called from the HTML service output. It runs the converter
 * on the active document.
 */
function createDownloadLink(syntax) {

  const document = DocumentApp.getActiveDocument();

  const name = document.getName() + '-' + timestamp();
  const data = syntax === 'xml'
    ? convertDocument(document)
    : document.getBody().getText();
  const ext = syntax === 'xml'
    ? 'xml'
    : 'txt';

  return { name, data, ext };
}

function testFunction() {
  var document = DocumentApp.getActiveDocument();
  Logger.log(convertDocument(document));
}

function convertDocument(document) {

  const body = document.getBody();
  const numChildren = body.getNumChildren();

  // used to track the foreground and background colours in use
  const fgColours = [];
  const bgColours = [];

  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<document>';

  // entry data
  let entries = [];
  let currEntryXml = '';

  // list data
  let listIndex = 1;
  let listId = null;
  let listIndices = {};
  
  for (let i = 0; i < numChildren; i++) {

    const child = body.getChild(i);

    /*
    if (isListItem(child)) {
      xml += `<number>${listIndex++}</number>\n`;
    } else if (isH1(child)) {
      xml += `<year>${child.getText()}</year>\n`;
    } else {
      let n = child.getNumChildren();
      for (j = 0; j < n; j++) {
        xml += `<p>${child.getChild(j).getText()}</p>\n`;
      }
    }
    */

    if (isListItem(child)) {

      if (!listId) {
        listId = child.getListId();
      }

      // add number
      if (listId == child.getListId()) {

        // dump out anything in currEntryXml and add to entries
        entries.push(adjustEntryXml(currEntryXml));

        // add the number
        currEntryXml = `<number>${listIndex++}</number>\n`;

        // convert the rest of the text
        currEntryXml += convertListItem(child) + '\n';
      } else {
        let currListId = child.getListId();
        let currListIndex = listIndices.hasOwnProperty(currListId)
          ? listIndices[currListId]
          : 1;
        // add the number to the start of the list item
        currEntryXml += convertListItem(child)
          .replace(/^(<body>)/g, `$1${currListIndex}. `);
        currEntryXml += '\n';
        listIndices[currListId] = ++currListIndex;
      }

    } else if (isH1(child)) {

      // dump out anything in currXmlPart and add to xmlParts
      currEntryXml = adjustEntryXml(currEntryXml);

      if (currEntryXml !== '') {
        entries.push(adjustEntryXml(currEntryXml));
      }

      // convert this h1
      currEntryXml = `<year>${child.getText()}</year>`;

    } else {
      currEntryXml += convertListItem(child) + '\n';
    }
  }

  Logger.log('FOREGROUND COLOURS:');
  Logger.log(fgColours);
  Logger.log('BACKGROUND COLOURS:');
  Logger.log(bgColours);

  // add the previous xmlPart
  entries.push(adjustEntryXml(currEntryXml));

  // join all of the entries together
  xml += entries.join('\n');

  // close the document
  xml += '</document>';

  return xml;

  function adjustEntryXml(entryXml) {
    return entryXml
      .trim()
      // remove trailing whitespace after block-level closing tags:
      // - </title>
      // - </subtitle>
      // - </body>
      // - </annotation>
      .replace(/<\/title>([ \f\t\v]+?)$/gm, '</title>')
      .replace(/<\/subtitle>([ \f\t\v]+?)$/gm, '</subtitle>')
      .replace(/<\/body>([ \f\t\v]+?)$/gm, '</body>')
      .replace(/<\/annotation>([ \f\t\v]+?)$/gm, '</annotation>')
      // remove extra newlines between <title> and <subtitle> nodes
      .replace(/(<\/title>)(\s+)(<subtitle>)/g, '$1\n$3')
      // remove extra newlines between <subtitle> and <body> nodes
      .replace(/(<\/subtitle>)(\s+)(<body>)/g, '$1\n$3')
      // replace valid empty lines with <body> nodes
      .replace(/(<\/body>\n)(\s+?)(<body>)/g, '$1<body></body>\n$3')
      // replace empty lines between <body> and <annotation> nodes with a
      // <body> node
      .replace(/(<\/body>\n)(\s*?)(<annotation>)/g, '$1<body></body>\n$3')
      // replace newlines at the start of body nodes
      .replace(/(<body>)\n/, '$1');
  }

  function convertListItem(listItem) {

    const numChildren = listItem.getNumChildren();
    let nodes = [];
    
    for (let i = 0; i < numChildren; i++) {
  
      const child = listItem.getChild(i);
  
      if (isText(child) && !isEmpty(child)) {
          const textStr = convertInlineElements(child)
            // replace carriage returns with newlines
            .replace(/\r/g, '\n')
            // remove empty lines between tags
            .replace(/>\n*\n</g, '>\n<');
          nodes = nodes.concat(textStr);
      } else {
        Logger.log(`Could not convert list item child: ${child.getType()}`);
      }
    }
    
    // if the paragraph is empty (no child nodes), a newline will be returned.
    // return nodes.join('\n') + '\n';
    // var converted = nodes.join('\n');
    // return converted;
    if (isListItem(listItem)) {
      return nodes.join('\n').trim();
    } else {
      return nodes.join('\n');
    }
  }

  function convertH1(h1) {

    const numChildren = h1.getNumChildren();
    let xml = '';

    for (let i = 0; i < numChildren; i++) {
      const child = h1.getChild(i);
      xml += `<h1>${child.getText()}</h1>\n`;
    }

    return xml;
  }

  function convertInlineElements(textElem) {  

    const indices = textElem.getTextAttributeIndices();
    let attsArray = [];
    
    for (const index of indices) {

      let atts = textElem.getAttributes(index);
      atts.INDEX = index;
      atts.IN = [];
      atts.OUT = [];
      attsArray.push(atts);

      // keep track of the foreground / background colours in use
      if (atts.FOREGROUND_COLOR && !fgColours.includes(atts.FOREGROUND_COLOR)) {
        fgColours.push(atts.FOREGROUND_COLOR);
      }
      if (atts.BACKGROUND_COLOR && !bgColours.includes(atts.BACKGROUND_COLOR)) {
        bgColours.push(atts.BACKGROUND_COLOR);
      }
    }

    // add an element for the very end of the text
    let endAtts = {
      INDEX: (textElem.getText().length),
      IN: [], // technically there would be no 'IN' attributes here, but . . . 
      OUT: []
    };
    ATTRIBUTES.forEach(att => {
      endAtts[att] = null;
    });
    attsArray.push(endAtts);

    try {
      attsArray = handleFormatting(attsArray);
      attsArray = addInsAndOuts(attsArray);
      attsArray = removeEmptyRuns(attsArray);
      attsArray = trimWhitespace(attsArray, textElem);
      attsArray = sortInsAndOuts(attsArray);
      return addSyntax(attsArray, textElem);
    } catch (error) {
      Logger.log(`ERROR: ${error.message}`);
      return '';
    }
  }
}

/**
 * Initiates all the functions below, returning a fully markdown-itized string
 * for the paragraph.
 */

function handleFormatting(attsArray) {

  attsArray.forEach(item => {

    for (const { name, style } of FORMATTING) {

      const {
        fgColours,
        bgColours
      } = style;

      const fgMatch = fgColours != null
        ? item.FOREGROUND_COLOR && fgColours.includes(item.FOREGROUND_COLOR)
        : true;

      // const bgMatch = bgColours != null
      //   ? item.BACKGROUND_COLOR && bgColours.includes(item.BACKGROUND_COLOR)
      //   : true;
      const bgMatch = bgColours == null
        ? true
        : bgColours.includes(null)
          ? bgColours.includes(item.BACKGROUND_COLOR)
          : item.BACKGROUND_COLOR && bgColours.includes(item.BACKGROUND_COLOR);

      item[name] = fgMatch && bgMatch;
    }
  });

  return attsArray;
}

/**
 * Removes the underline attribute from links and creates a 'LINK' attribute
 * with a boolean (or null) value for all other elements.
 */
 function handleLinks(attsArray) {

  attsArray.forEach(function (atts) {
    if (atts['LINK_URL'] !== null) {
      atts['LINK'] = true;
      atts['UNDERLINE'] = null;
    } else {
      atts['LINK'] = null;
    }
  });

  return attsArray;
}

/** 
 * Adds 'IN' and 'OUT' arrays to the array of attributes to identify which
 * styles (of the ones we care about) are beginning or ending at a given point.
 */
function addInsAndOuts(arr) {

  if (arr.length === 0) {
    return arr;
  }

  // the attributes at the start can only be 'IN' attributes
  ATTRIBUTES.forEach(function (attr) {
    if (arr[0][attr]) {
      arr[0].IN.push(attr);
    }
  });

  for (let i = 1; i < arr.length; i++) {

    ATTRIBUTES.forEach(function (attr) {

      const prev = arr[i - 1][attr];
      const curr = arr[i][attr];
      
      // add in attributes
      if (curr && !prev) {
        arr[i].IN.push(attr);
      }

      // and now all the out ones.
      if (prev && !curr) {
        arr[i].OUT.push(attr);
      }

    });
  }

  return arr;
}

function removeEmptyRuns(attsArray) {

  var newAttsArray = JSON.parse(JSON.stringify(attsArray));

  // Remove elements we don't care about.
  for (let i = 0; i < newAttsArray.length; i++) {
    const atts = newAttsArray[i];
    if ((atts.IN.length === 0) && (atts.OUT.length === 0))  {
      newAttsArray.splice(i, 1);
    }
  };

  return newAttsArray;
}

function trimWhitespace(attsArray, textElem) {

  let newAttsArray = JSON.parse(JSON.stringify(attsArray));
  const textStr = textElem.getText();

  // Avoid leading whitespace inside style tags.
  let staticLength = newAttsArray.length - 1;
  for (let i = 0; i < staticLength; i++) {

    let currObj = newAttsArray[i];
    let nextObj = newAttsArray[i+1];
    let currText = textStr.slice(currObj.INDEX, nextObj.INDEX);

    if (currObj.IN.length && /^\s+/.test(currText)) {
      const offset = /^\s+/.exec(currText)[0].length;
      if (currObj.OUT.length === 0) {
        currObj.INDEX = currObj.INDEX + offset;
      } else {
        newAttsArray.push({
          ...currObj,
          INDEX: currObj.INDEX + offset,
          OUT: []
        });
        currObj.IN = [];
      }
    }
  }

  // reorder the array so that trailing whitespace can be handled correctly
  newAttsArray.sort((a, b) => a.INDEX - b.INDEX);
  newAttsArray = consolidateArray(newAttsArray);

  // Avoid trailing whitespace inside style tags.
  staticLength = newAttsArray.length;
  for (let i = 1; i < staticLength; i++) {

    let prevObj = newAttsArray[i-1];
    let currObj = newAttsArray[i];
    let prevText = textStr.slice(prevObj.INDEX, currObj.INDEX);

    if (currObj.OUT.length && /\s+$/.test(prevText)) {
      const offset = /\s+$/.exec(prevText)[0].length;

      // if the IN array is empty, we can repurpose this atts object for this
      // new style range. otherwise, create a brand new style range, 
      // effectively splitting this style range in two.
      if (currObj.IN.length === 0) {
        currObj.INDEX = currObj.INDEX - offset;
      } else {
        newAttsArray.push({
          ...currObj,
          INDEX: currObj.INDEX - offset,
          IN: []
        });
        currObj.OUT = [];
      }
    }
  }

  newAttsArray.sort((a, b) => a.INDEX - b.INDEX);
  newAttsArray = consolidateArray(newAttsArray);

  return newAttsArray;
}

function consolidateArray(attsArr) {

  if (attsArr.length === 0) {
    return attsArr;
  }

  let newAttsArr = [];
  let prevAtts = attsArr[0];

  for (let i = 1; i < attsArr.length; i++) {

    let currAtts = attsArr[i];

    if (currAtts.INDEX !== prevAtts.INDEX) {
      newAttsArr.push(prevAtts);
      prevAtts = currAtts;
      continue;
    } else {

      // merge the current item with the item held in memory, so to speak
      const IN = new Set([
        ...prevAtts.IN,
        ...currAtts.IN
      ]);
      const OUT = new Set([
        ...prevAtts.OUT,
        ...currAtts.OUT
      ]);

      prevAtts = {
        ...prevAtts,
        ...currAtts,
        // we want the set difference, because something being in both the IN
        // and OUT array is meaningless -- no need to stop then start again.
        IN: [...setDifference(IN, OUT)],
        OUT: [...setDifference(OUT, IN)],
      };
    }
  }

  newAttsArr.push(prevAtts);

  return newAttsArr;
}

/**
 * Sort the IN and OUT arrays such that they are mirror images of each other,
 * of sorts. That is, the IN array should follow the order of ATTRIBUTES and
 * the OUT array should be in the reverse order of ATTRIBUTES, so close tags
 * mirror open tags properly.
 * @param {*} attsArray 
 * @returns 
 */
function sortInsAndOuts(attsArray) {

  for (let i = 0; i < attsArray.length; i++) {
    attsArray[i].IN.sort((a, b) => (
      ATTRIBUTES.indexOf(a) - ATTRIBUTES.indexOf(b)
    ));
    attsArray[i].OUT.sort((a, b) => (
      ATTRIBUTES.indexOf(b) - ATTRIBUTES.indexOf(a)
    ));
  }

  return attsArray;
}

/**
 * Creates a new string with markdown syntax based on the array of style
 * inflection points.
 * 
 * Has a number of other iterators:
 *   - removes elements from the array with empty IN or OUT arrays
 *   - moves inflection points to avoid 'styled' whitespace
 *   - sorts the [IN] and [OUT] arrays to properly handle nested styles.
 */
 function addSyntax(attsArr, textElem){

  const textStr = textElem.getText();
  let currIndex = 0;
  let textWithSyntax = '';
  
  for (const atts of attsArr) {
    
    // add the text from the previous run
    if (currIndex < atts.INDEX) {
      const substr = textStr.substring(currIndex, atts.INDEX)
      textWithSyntax += entitize(substr);
      currIndex = atts.INDEX;
    }

    // Outgoing styles should be added first, in the unusual event that a style
    // stops exactly where the next one begins.
    for (const att of atts.OUT) {
      textWithSyntax += MAP[att].end;
    }
    
    for (const att of atts.IN) {
      textWithSyntax += MAP[att].start;
    }
  }

  // if there's still text at the end, add it
  // what about the OUT array at the end?
  const endAtts = attsArr[attsArr.length - 1];
  if (textStr.length > endAtts.INDEX) {
    const substr = textStr.substring(endAtts.INDEX);
    textWithSyntax += entitize(substr);
  }

  return textWithSyntax;  
}

function entitize(str) {
  return str.replace(/&/g, '&#38;')
    .replace(/</g, '&#60;')
    .replace(/>/g, '&#62;');
}
