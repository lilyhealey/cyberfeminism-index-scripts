var MAIN_INDEX_NUMBER_PS_NAME = 'number';
var MAIN_INDEX_NUMBER_OS_NAME = 'main-index-number';
var CENTERED_PS_NAME = 'number';

// get the MAIN_INDEX_NUMBER_PS
function getParStyleByName(doc, parStyleName) {
  for (var i = 0; i < doc.allParagraphStyles.length; i++) {
    var ps = doc.allParagraphStyles[i];
    if (ps.name == parStyleName) {
      return ps;
    }
  }

  return null;
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

// scan scan all of the stories for text with MAIN_INDEX_NUMBER_PS par style

// if the parent text frame of MAIN_INDEX_NUMBER_PS is NOT
// MAIN_INDEX_NUMBER_OS, then:
// 1. create a new object
// 2. apply MAIN_INDEX_NUMBER_OS
// 3. insert text
// 4. cut object and paste where MAIN_INDEX_NUMBER text was

function createIndexNumberObject(doc, text, par) {
  var os = getObjectStyleByName(doc, MAIN_INDEX_NUMBER_OS_NAME);
  var tf = doc.textFrames.add();

  tf.contents = text;
  tf.appliedObjectStyle = os;
  tf.select();
  app.cut();

  // story.insertionPoints.lastItem().select();
  par.characters[0].select();
  for (var i = 1; i < par.characters.length - 1; i++) {
    par.characters[i].select(SelectionOptions.ADD_TO);
  }
  app.paste();
}

function findParagraphAndSelect(doc) {
  var mainIndexPs = getParStyleByName(doc, MAIN_INDEX_NUMBER_PS_NAME);
  var centeredPs = getParStyleByName(doc, CENTERED_PS_NAME);
  var numStories = doc.stories.length;

  for (var k = 0; k < numStories; k++) {
    var story = doc.stories[k];
    for (var i = 0; i < story.paragraphs.length; i++) {
      var par = story.paragraphs[i];
      if (par.appliedParagraphStyle == mainIndexPs) {
        var ignore = false;
        for (var j = 0; j < par.parentTextFrames.length; j++) {
          if (par.parentTextFrames[j].appliedObjectStyle.name == MAIN_INDEX_NUMBER_OS_NAME) {
            ignore = true;
            break;
          }
        }
        if (!ignore) {
          par.appliedParagraphStyle = centeredPs;
          createIndexNumberObject(doc, par.contents, par);
        }
      }
    }
  }
}

function alignIndexNumbers(doc) {
  for (var k = 0; k < doc.stories.length; k++) {
    var story = doc.stories[k];
    for (var j = 0; j < story.paragraphs.length; j++) {
      var graf = story.paragraphs[j];
      if (graf) {
        for (var i = 0; i < graf.textFrames.length; i++) {
          var tf = graf.textFrames[i];
          doc.align([tf], AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        }
      }
    }
  }
  // alert(doc.textFrames.length);
  // alert(doc.textFrames[i].textFrames.length);
}
// get the active document
function main() {

  var doc = app.activeDocument;
  findParagraphAndSelect(doc);
  alignIndexNumbers(doc);
  alert('done');
  // createIndexNumberObject(doc, '3');
}

main();

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
