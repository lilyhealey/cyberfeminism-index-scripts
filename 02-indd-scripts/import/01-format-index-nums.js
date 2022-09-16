var ENTRY_NUMBER_OS_NAME = 'CR - IN - entry header';
var RAW_TAG_NAME = 'number';
var FORMATTED_TAG_NAME = 'formatted_number';

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

function createProgressBarWindow(args) {
  // parse out the args
  var title = args.title || 'Progress';
  var maximumValue = args.maximumValue;
  var width = args.width;

  var progressBarPanel = new Window(
    // type: 'window' | 'palette' | 'dialog'
    'window',
    // title
    title,
    // bounds
    undefined,
    // properties
    { borderless: true }
  );

  progressBarPanel.progressBar = progressBarPanel.add(
    'progressbar',
    [12, 12, width, 24],
    0,
    maximumValue
  );

  return progressBarPanel;
}

function getObjectStyleByName(doc, objectStyleName) {
  for (var i = 0; i < doc.allObjectStyles.length; i++) {
    var os = doc.allObjectStyles[i];
    if (os.name == objectStyleName) {
      return os;
    }
  }
  return null;
}

function getParagraphStyleByName(doc, parStyleName) {
  for (var i = 0; i < doc.allParagraphStyles.length; i++) {
    var ps = doc.allParagraphStyles[i];
    if (ps.name == parStyleName) {
      return ps;
    }
  }
  return null;
}

function createIndexNumberObject(doc, contents, text) {
  var os = getObjectStyleByName(doc, ENTRY_NUMBER_OS_NAME);
  var tf = doc.textFrames.add();

  tf.contents = contents;
  tf.appliedObjectStyle = os;
  tf.select();
  app.cut();

  // story.insertionPoints.lastItem().select();
  // text.characters[0].select();
  // for (var i = 1; i < text.characters.length; i++) {
  //   text.characters[i].select(SelectionOptions.ADD_TO);
  // }
  text.select();
  app.paste();
  
}

function main() {
  // get things ready for the main action
  app.scriptPreferences.enableRedraw = false;

  // find all of the numbers
  var doc = app.activeDocument;
  var root = doc.xmlElements[0];
  var numberElements = root.evaluateXPathExpression(RAW_TAG_NAME);
  var numNumbers = numberElements.length;

  // create the progress bar
  var progressBarPanel = createProgressBarWindow({
    maximumValue: numNumbers,
    width: 400,
    title: 'Formatting Index Entry Numbers'
  });
  progressBarPanel.show();
  var ps = getParagraphStyleByName(doc, '10/12 - IN - left rag');

  // cycle over all of the numbers
  for (var i = numNumbers - 1; i >= 0; i--) {
    progressBarPanel.progressBar.value = numNumbers - i;
    var el = numberElements[i];
    var text = el.xmlContent;
    createIndexNumberObject(doc, el.contents, text);
    el.markupTag = FORMATTED_TAG_NAME;
    
    var ip = text.insertionPoints.firstItem();
    ip.contents += '\r\r';
    ip.appliedParagraphStyle = ps;
    
    ip = text.insertionPoints.lastItem();
    ip.contents += '\r';
  }

  // cleanup
  progressBarPanel.close();
  app.scriptPreferences.enableRedraw = true;
  alert('done');
}

main();
