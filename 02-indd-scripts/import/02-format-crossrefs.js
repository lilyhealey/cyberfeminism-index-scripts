var ENTRY_CROSSREF_OS_NAME = 'CR - IN - entry text';
var RAW_TAG_NAME = 'crossref';
var FORMATTED_TAG_NAME = 'formatted_crossref';

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

  var progressBarWindow = new Window(
    // type: 'window' | 'palette' | 'dialog'
    'window',
    // title
    title,
    // bounds
    undefined,
    // properties
    { borderless: true }
  );

  progressBarWindow.progressBar = progressBarWindow.add(
    'progressbar',
    [12, 12, width, 24],
    0,
    maximumValue
  );

  return progressBarWindow;
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

function createIndexNumberObject(doc, contents, text) {
  var os = getObjectStyleByName(doc, ENTRY_CROSSREF_OS_NAME);
  var textFrame = doc.textFrames.add();
  textFrame.contents = contents;
  textFrame.appliedObjectStyle = os;
  textFrame.select();
  app.cut();
  text.select();
  app.paste();
}

function main() {
  // get things ready for the main action
  app.scriptPreferences.enableRedraw = false;

  // find all of the numbers
  var doc = app.activeDocument;
  var root = doc.xmlElements[0];
  var numberElements = root.evaluateXPathExpression('//' + RAW_TAG_NAME);
  var numNumbers = numberElements.length;

  // create the progress bar
  var progressBarWindow = createProgressBarWindow({
    maximumValue: numNumbers,
    width: 400,
    title: 'Formatting Inline Index Numbers'
  });
  progressBarWindow.show();

  // cycle over all of the numbers
  for (var i = numNumbers - 1; i >= 0; i--) {
    // update the progress bar
    progressBarWindow.progressBar.value = numNumbers - i;

    // 
    var el = numberElements[i];
    var text = el.xmlContent;

    // remove the parens around the numbers
    var contents = el.contents
      .replace(/\(/g, '')
      .replace(/\)/g, '');

    // create the anchored object with object style
    createIndexNumberObject(doc, contents, text);

    // mark the element as 'done' by changing the kind of tag
    el.markupTag = FORMATTED_TAG_NAME;
  }

  // clean up
  progressBarWindow.close();
  app.scriptPreferences.enableRedraw = true;
  alert('done');
}

main();
