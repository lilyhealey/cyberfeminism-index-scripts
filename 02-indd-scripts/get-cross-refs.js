var CS = 'AS - green';
var OS = 'Entry Crossref';
var SIDENOTES_OS = 'sidenotes';
var SIDENOTE_NUMBER_OS_NAME = 'Sidenote Number';
var BOTTOM_WRAP = 5;
var STARTING_WIDTH = 50;
var INCREMENT = 1;

function createIndexNumberObject(spread, text, sidenoteFrame) {
  var doc = app.activeDocument;
  var os = getObjectStyleByName(doc, SIDENOTE_NUMBER_OS_NAME);
  var tf = spread.textFrames.add();

  tf.contents = text.toString();
  tf.appliedObjectStyle = os;
  tf.select();
  app.cut();

  // story.insertionPoints.lastItem().select();
  // par.characters[0].select();
  // for (var i = 1; i < par.characters.length - 1; i++) {
  //   par.characters[i].select(SelectionOptions.ADD_TO);
  // }
  sidenoteFrame.insertionPoints.lastItem().select();
  app.paste();
}

// get the MAIN_INDEX_NUMBER_OS
function getObjectStyleByName(doc, objectStyleName) {
  for (var i = 0; i < doc.allObjectStyles.length; i++) {
    var os = doc.allObjectStyles[i];
    if (os.name == objectStyleName) {
      return os;
    }
  }

  return null;
}


function isCrossref(textStyleRange) {
  var appliedCS = textStyleRange.appliedCharacterStyle;
  return (
    (appliedCS.constructor.name === 'CharacterStyle' && appliedCS.name === CS)
    ||
    (appliedCS.constructor.name === 'String' && appliedCS === CS)
  );
}

function getCrossrefNumsOnSpread(spread) {
  var crossrefs = [];
  for (var i = 0; i < spread.textFrames.length; i++) {
    var tf = spread.textFrames[i];
    for (var j = 0; j < tf.textStyleRanges.length; j++) {
      var tsr = tf.textStyleRanges[j];
      if (isCrossref(tsr)) {
        var cr = tsr.contents.replace(/\D/g, '');
        crossrefs.push(parseInt(cr));
      }
    }
  }
  return crossrefs;
}

function getObjectCrossrefs(spread) {
  var crossrefs = [];
  for (var i = 0; i < spread.textFrames.length; i++) {
    var tf = spread.textFrames[i];
    for (var j = 0; j < tf.textFrames.length; j++) {
      var ttf = tf.textFrames[j];
      if (ttf.appliedObjectStyle && ttf.appliedObjectStyle.name == OS) {
        crossrefs.push(parseInt(ttf.contents));
      }
    }
  }
  return crossrefs;
}

function getTitles() {
  var doc = app.activeDocument;
  var root = doc.xmlElements[0];
  var titleElements = root.evaluateXPathExpression('title');

  var titles = [];
  for (var i = titleElements.length - 1; i >= 0; i--) {
    titles.push(titleElements[i].contents);
  }

  return titles;
}

function getMargins() {
  var spread = app.activeWindow.activeSpread;
  for (var i = 0; i < spread.pages.length; i++) {
    alert([
      spread.pages[i].marginPreferences.top,
      spread.pages[i].marginPreferences.right,
      spread.pages[i].marginPreferences.bottom,
      spread.pages[i].marginPreferences.left,
    ].join(', '));
  }
}

function getSidenotesObjecStyle() {
  var doc = app.activeDocument;
  for (var i = 0; i < doc.allObjectStyles.length; i++) {
    var os = doc.allObjectStyles[i];
    if (os.name === SIDENOTES_OS) {
      return os;
    }
  }
  return null;
}

function addTextFrame(contents, numsAndTitles) {
  var spread = app.activeWindow.activeSpread;
  var recto = spread.pages[1];
  var y1 = recto.bounds[0] + recto.marginPreferences.top;
  var y2 = recto.bounds[2] - recto.marginPreferences.bottom;
  var x2 = recto.bounds[3] - recto.marginPreferences.right;
  var x1 = x2 - STARTING_WIDTH;
  var tf = recto.textFrames.add(undefined, undefined, undefined, {
    geometricBounds: [0, 0, 10, 10],
    absoluteRotationAngle: -90,
  });

  tf.appliedObjectStyle = getSidenotesObjecStyle();
  var gb = [y1, x1, y2, x2];
  tf.geometricBounds = gb;

  for (var i = 0; i < numsAndTitles.length; i++) {
    var num = numsAndTitles[i].num;
    var title = numsAndTitles[i].title;
    createIndexNumberObject(spread, num, tf);
    tf.insertionPoints.lastItem().contents += ' ';
    tf.insertionPoints.lastItem().contents += title;
    tf.insertionPoints.lastItem().contents += '\r';
  }
  // tf.contents = contents;

  // gb[1] = gb[1] + 10;
  // tf.geometricBounds = gb;

  var i = 0;
  do {
    i++;
    if (i > 100) {
      break;
    }
    gb[1] = gb[1] + INCREMENT; 
    tf.geometricBounds = gb;
  } while (!tf.overflows);

  gb[1] = gb[1] - INCREMENT;
  tf.geometricBounds = gb;

  // format all of the numbers
}

/**
 * Writes to a log.
 * 
 * @param {string} s - the log to write
 */
 function boop(s) {

  var now = new Date();
  var output = now.toLocaleTimeString() + ": " + s;
  var logFolder = Folder(Folder.desktop.fsName + "/Logs");
  if (!logFolder.exists) {
    logFolder.create();
  }

  var logFile = File(logFolder.fsName + "/indesign_log.txt");
  logFile.lineFeed = "Unix";
  logFile.encoding = "UTF-8";
  logFile.open("a");
  logFile.writeln(output);
  logFile.close();

}

function main() {
  var spread = app.activeWindow.activeSpread;
  // var crossrefNums = getCrossrefNumsOnSpread(spread);
  var crossrefNums = getObjectCrossrefs(spread);
  var titles = getTitles();
  var s = '';
  var numsAndTitles = [];
  for (var i = 0; i < crossrefNums.length; i++) {
    var num = crossrefNums[i];
    var title = titles[num-1];
    numsAndTitles.push({
      num: num,
      title: title
    });
    if (s.length > 0) {
      s += '\r';
    }
    s += '(' + num.toString() + ') ' + title;
  }
  addTextFrame(s, numsAndTitles);
  alert('done');
}

app.scriptPreferences.enableRedraw = false;
main();
app.scriptPreferences.enableRedraw = true;
