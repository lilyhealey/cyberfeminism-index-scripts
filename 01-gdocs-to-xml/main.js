const VERBOSE = false;
const map = {
  syntax: 'markup',
  ext: 'xml',
  inline: {
    italic:     { start: '<italic>',    end: '</italic>'  },
    bold:       { start: '<bold>',       end: '</bold>'    },
    crossref:   { start: '<crossref>',   end: '</crossref>'},
    annotation: { start: '<annotation>', end: '</annotation>\n'},
    blockcrossref: { start: '<!-- ', end: ' -->\n'},
    link: {
      start: (url) => `<a href="${url}">`,
      end:   () => '</a>',
    }
  }
};

/** 
 * Adds the menu item on open. 
 */
function onOpen(e) {
  const ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('Convert to XML', 'starterXML')
    .addToUi();
}

/** 
 * Returns an HTML dialog with the download link.
 */
function starterXML() {
  const template = HtmlService.createTemplateFromFile('dialog');
  const html = template.evaluate();
  html.setWidth(300)
    .setHeight(150); 
  DocumentApp.getUi().showModalDialog(html, 'XML Conversion');
}

/** 
 * This function is called from the HTML service output. It runs the converter
 * on the active document.
 */
function createDownloadLink() {
  const document = DocumentApp.getActiveDocument();
  Logger.log(document.getId());
  const name = DocumentApp.getActiveDocument().getName() + '_' + timestamp();
  const data = convertDocument(document);
  const ext = 'xml';
  return { name, data, ext };
}

function testFunction() {
  var document = DocumentApp.getActiveDocument();
  Logger.log(convertDocument(document));
}

function convertDocument(document) {

  const body = document.getBody();
  const numChildren = body.getNumChildren();

  const fgColours = [];
  const bgColours = [];
  let xml = '';
  let listIndex = 1;
  
  for (let i = 0; i < numChildren; i++) {

    const child = body.getChild(i);

    if (isListItem(child)) {

      // add number
      xml += `<number>${listIndex++}</number>\n`;

      xml += convertListItem(child);

    } else if (VERBOSE) {
      Logger.log('NOT PROCESSED:');
      Logger.log(child.getText());
    }
  }

  function convertListItem(listItem) {

    const numChildren = listItem.getNumChildren();
    let nodes = [];
    
    for (let i = 0; i < numChildren; i++) {
  
      const child = listItem.getChild(i);
  
      if (isText(child) && !isEmpty(child)) {
        const textStr = convertInlineElements(child);
        nodes = nodes.concat(convertBlockElements(textStr));
      } else if (VERBOSE) {
        Logger.log(child.getType());
      }
    }
    
    // if the paragraph is empty (no child nodes), a newline will be returned.
    // return nodes.join('\n') + '\n';
    var converted = nodes.join('\n') + '\n';
    Logger.log(converted);
    return converted;
  }

  function convertInlineElements(textElem) {  

    const indices = textElem.getTextAttributeIndices();
    let attsArray = [];
    
    for (const index of indices) {

      let atts = textElem.getAttributes(index);
      atts.INDEX = index;
      attsArray.push(atts);

      // keep track of the foreground / background colours in use
      if (atts.FOREGROUND_COLOR && !fgColours.includes(atts.FOREGROUND_COLOR)) {
        fgColours.push(atts.FOREGROUND_COLOR);
      }
      if (atts.BACKGROUND_COLOR && !bgColours.includes(atts.BACKGROUND_COLOR)) {
        bgColours.push(atts.BACKGROUND_COLOR);
      }
    }

    try {
      attsArray = handleFormatting(attsArray);
      attsArray = addInsAndOuts(attsArray, textElem);
      attsArray = removeEmptyRuns(attsArray);
      attsArray = trimWhitespace(attsArray, textElem);
      attsArray = sortInsAndOuts(attsArray);
      // Logger.log(JSON.stringify(attrArray, null, "  "));
      return addSyntax(attsArray, textElem, map);
    } catch (error) {
      Logger.log(`ERROR: ${error.message}`);
      return '';
    }
  }

  function convertBlockElements(text) {

    let nodes = [];

    // convert block-level elements
    var regex = /\r{2}/gi;
    if (regex.test(text)) {
      var parts = text.split(regex);
      var [header, ...rest] = parts;
      var headerParts = header.split(/\r+/g);
      for (var j = 0; j < headerParts.length; j++) {
        const headerNode = `<header>${headerParts[j]}</header>`;
        nodes.push(headerNode);
      }
      
      for (var j = 0; j < rest.length; j++) {
        const restPart = rest[j];
        if (!/^<(annotation>|!--)/gi.test(restPart)) {
          const restPartParts = restPart.split(/\r/g);
          for (const rpp of restPartParts) {
            const restPartNode = `<body>${rpp}</body>`;
            nodes.push(restPartNode);
          }
        } else {
          nodes.push(restPart);
        }
      }

    } else if (VERBOSE) {
      Logger.log('NO DOUBLE RETURN');
      Logger.log(text);
    }

    return nodes;
  }

  Logger.log('FOREGROUND COLOURS:');
  Logger.log(fgColours);
  Logger.log('BACKGROUND COLOURS:');
  Logger.log(bgColours);

  return xml;
}

/**
 * Initiates all the functions below, returning a fully markdown-itized string
 * for the paragraph.
 */

function handleFormatting(attrArray) {

  const arr = JSON.parse(JSON.stringify(attrArray));

  const formatting = [
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
      name: 'BLOCK_CROSSREF',
      style: {
        fgColours: null,
        bgColours: [
          '#d9d9d9'
        ]
      }
    },
    {
      name: 'IMAGE',
      style: {
        fgColours: null,
        bgColours: [
          '#d9ead3'
        ]
      }
    }
  ];

  arr.forEach(item => {

    for (const { name, style } of formatting) {

      const {
        fgColours,
        bgColours
      } = style;


      const fgMatch = fgColours != null
        ? item.FOREGROUND_COLOR && fgColours.includes(item.FOREGROUND_COLOR)
        : true;

      const bgMatch = bgColours != null
        ? item.BACKGROUND_COLOR && bgColours.includes(item.BACKGROUND_COLOR)
        : true;

      item[name] = fgMatch && bgMatch;
    }
  });

  return arr;
}

/**
 * Removes the underline attribute from links and creates a 'LINK' attribute
 * with a boolean (or null) value for all other elements.
 */
 function handleLinks(attrArray) {

  var linkArr = JSON.parse(JSON.stringify(attrArray));

  linkArr.forEach(function (item) {
    if (item['LINK_URL'] !== null) {
      item['LINK'] = true;
      item['UNDERLINE'] = null;
    } else {
      item['LINK'] = null;
    }
  });

  return linkArr;
}

/** 
 * Adds 'IN' and 'OUT' arrays to the array of attributes to identify which
 * styles (of the ones we care about) are beginning or ending at a given point.
 */
function addInsAndOuts(array, textElem) {

  var inOutArr = JSON.parse(JSON.stringify(array));
  var attrs = [
    'ITALIC',
    'LINK',
    'BOLD',
    'UNDERLINE',
    'CROSSREF',
    'BLOCK_CROSSREF',
    'ANNOTATION',
    'IMAGE'
  ];
  
  //add an element for the very end of the text
  inOutArr.push({
    INDEX: (textElem.getText().length), 
    ITALIC: null, 
    LINK: null, 
    BOLD: null, 
    UNDERLINE: null,
    CROSSREF: null,
    ANNOTATION: null,
    BLOCK_CROSSREF: null,
    IMAGE: null,
    IN: [], 
    OUT: []
  });
  
  var staticLength = inOutArr.length;
  
  for (var i = 0; i < staticLength; i++) {

    var obj = inOutArr[i];
    delete obj.STRIKETHROUGH;
    obj.IN = [];
    obj.OUT = [];
    
    // The first attribute index has to be an 'in.' 
    if (i === 0) {
      attrs.forEach(function (attr) {
        if (obj[attr]) {
          obj.IN = [];
          obj.IN.push(attr);
        }
      });
    } else {

      attrs.forEach(function (attr) {
        if (obj[attr] && !(inOutArr[i - 1][attr])) {
          obj.IN.push(attr);
        }
      });

      // and now all the out ones.
      attrs.forEach(function (attr) {
        if (inOutArr[i - 1][attr] && !(obj[attr])) {
          obj.OUT.push(attr);
        }
      });
    } 
  }
  
  return inOutArr;
}

function removeEmptyRuns(array) {

  var newArray = JSON.parse(JSON.stringify(array));

  // Remove elements we don't care about.
  for (var i = 0; i < newArray.length; i++) {
    if ((newArray[i].IN.length === 0) && (newArray[i].OUT.length === 0))  {
      newArray.splice(i, 1) 
    }
  };

  return newArray;
}

function trimWhitespace(array, textElem) {

  var theText = textElem.getText();
  var newArray = JSON.parse(JSON.stringify(array));
  var staticLength = newArray.length - 1;

  // Avoid leading whitespace inside style tags.
  for (var i = 1; i < staticLength; i++) {

    var prevObj = newArray[i-1];
    var currObj = newArray[i];
    var nextObj = newArray[i+1];

    var prevText = theText.slice(prevObj.INDEX, currObj.INDEX);
    var currText = theText.slice(currObj.INDEX, nextObj.INDEX);

    if (currObj.IN.length && /^\s+/.test(currText)) {
      const result = /^\s+/.exec(currText);
      if (currObj.OUT.length === 0) {
        currObj.INDEX = currObj.INDEX + result[0].length;
      } else {
        newArray.push({
          ...currObj,
          INDEX: currObj.INDEX + result[0].length,
          OUT: []
        });
        currObj.IN = [];
      }
    }
  }

  newArray.sort((a, b) => a.INDEX - b.INDEX);
  newArray = consolidateArray(newArray);
  staticLength = newArray.length - 1;

  // Avoid trailing whitespace inside style tags.
  for (var i = 1; i < staticLength; i++) {

    var prevObj = newArray[i-1];
    var currObj = newArray[i];
    var nextObj = newArray[i+1];

    var prevText = theText.slice(prevObj.INDEX, currObj.INDEX);
    var currText = theText.slice(currObj.INDEX, nextObj.INDEX);

    if (currObj.OUT.length && /\s+$/.test(prevText)) {
      const result = /\s+$/.exec(prevText);
      if (currObj.IN.length === 0) {
        currObj.INDEX = currObj.INDEX - result[0].length;
      } else {
        newArray.push({
          ...currObj,
          INDEX: currObj.INDEX - result[0].length,
          IN: []
        });
        currObj.OUT = [];
      }
    }
  }

  newArray.sort((a, b) => a.INDEX - b.INDEX);
  newArray = consolidateArray(newArray);
  return newArray;
}

function consolidateArray(arr) {

  if (arr.length === 0) {
    return arr;
  }

  var newArray = [];
  var item = arr[0];

  for (var i = 1; i < arr.length; i++) {

    var currItem = arr[i];

    if (currItem.INDEX !== item.INDEX) {
      newArray.push(item);
      item = currItem;
      continue;
    } else {

      // merge the current item with the item held in memory, so to speak
      const IN = new Set([
        ...item.IN,
        ...currItem.IN
      ]);
      const OUT = new Set([
        ...item.OUT,
        ...currItem.OUT
      ]);

      item = {
        ...item,
        ...currItem,
        // we want the set difference, because something being in both the IN
        // and OUT array is meaningless -- no need to stop then start again.
        IN: [...setDifference(IN, OUT)],
        OUT: [...setDifference(OUT, IN)],
      };
    }

  }

  newArray.push(item);

  return newArray;
}

function sortInsAndOuts(attsArray) {

  for (let i = 0; i < attsArray.length; i++) {
    // The two sort functions below make sure that if two styles start or end
    // at the same place, they don't get put in the wrong order (the one that
    // starts first ends last, etc.)
    if (attsArray[i].IN.length > 1) {
      attsArray[i].IN.sort(
        function (attA, attB) {
          for (let j = 0; j < (attsArray.length - i) ; j++) {
            const indexA = attsArray[i + j].OUT.indexOf(attA);
            const indexB = attsArray[i + j].OUT.indexOf(attB);
            if (indexA > -1 || indexB === -1) {
              return -1;
            } else if (indexB > -1 || indexA === -1) {
              return 1;
            } else if (indexB > -1 || indexA > -1) {
              return 0; 
            }
          }
        }
      );
    }
  
    if (attsArray[i].OUT.length > 1) {
      attsArray[i].OUT.sort(
        function (attA, attB) {
          for (let k = 0; k < i; k++) {
            const indexA = attsArray[i - k].IN.indexOf(attA);
            const indexB = attsArray[i - k].IN.indexOf(attB);
            if (indexA > -1 || indexB === -1) {
              return 1;
            } else if (indexB > -1 || indexA === -1) {
              return -1;
            } else if (indexB > -1 || indexA > -1) {
              return 0; 
            }
          }
        }
      ); 
    }
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
 function addSyntax(array, textElem, map){
  
  const textStr = textElem.getText();
  let currIndex = 0;
  var attsArray = JSON.parse(JSON.stringify(array));
  var textWithSyntax = "";
  
  for (let i = 0; i < attsArray.length; i++) {
    
    // add the text from the previous run
    if (currIndex < attsArray[i].INDEX) {
      const subtext = textStr.substring(currIndex, attsArray[i].INDEX);
      // if (map.syntax === maps.markdown.syntax) {
      //   textWithSyntax += escapeLists(subtext);
      // }
      // if (map.syntax === maps.markup.syntax) {
      //   textWithSyntax += entitize(subtext);
      // }
      textWithSyntax += subtext;
      currIndex = attsArray[i].INDEX;
    }

    // Outgoing styles should be added first, in the unusual event that a style
    // stops exactly where the next one begins.
    for (const att of attsArray[i].OUT) {
      if (att === 'ITALIC') {
        textWithSyntax += map.inline.italic.end; 
      } else if (att === 'BOLD') {
        textWithSyntax += map.inline.bold.end;
      } else if (att === 'LINK') {
        var url = attsArray[i - 1].LINK_URL;
        textWithSyntax += (map.inline.link.end(url));
      } else if (att === 'ANNOTATION') {
        textWithSyntax += map.inline.annotation.end;
      } else if (att === 'CROSSREF') {
        textWithSyntax += map.inline.crossref.end;
      } else if (att === 'BLOCK_CROSSREF' || att === 'IMAGE') {
        textWithSyntax += map.inline.blockcrossref.end;
      }
    }
    
    for (const att of attsArray[i].IN) {
      if (att === 'ITALIC') {
        textWithSyntax += map.inline.italic.start; 
      } else if (att === 'BOLD') {
        textWithSyntax += map.inline.bold.start;
      } else if (att === 'LINK') {
        var url = attsArray[i].LINK_URL;
        textWithSyntax += (map.inline.link.start(url));
      } else if (att === 'ANNOTATION') {
        textWithSyntax += map.inline.annotation.start;
      } else if (att === 'CROSSREF') {
        textWithSyntax += map.inline.crossref.start;
      } else if (att === 'BLOCK_CROSSREF' || att === 'IMAGE') {
        textWithSyntax += map.inline.blockcrossref.start;
      }
    }
    
    // if we're at the end, and there is still text left, add it.
    if (i === (attsArray.length - 1) && textStr.length > attsArray[i].INDEX) {
      const subtext = textStr.substring(attsArray[i].INDEX);
      // if (map.syntax === maps.markdown.syntax) {
      //   textWithSyntax += escapeLists(subtext); 
      // }
      // if (map.syntax === maps.markup.syntax) {
      //   textWithSyntax += entitize(subtext);
      // }
      textWithSyntax += subtext;
    }

  }
  return textWithSyntax;  
}
