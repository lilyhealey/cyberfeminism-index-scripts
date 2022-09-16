// ============================================================================
// reapply-object-styles.js
//
// given a list of pre-defined object styles, scan all of the page items of
// the current indesign document for objects with one of styles applied, and
// reapply the style to clear overrides.
// ============================================================================

// every page item that currently has one of the below object styles applied
// will have the style re-applied, with overrides cleared.
var O_STYLE_NAMES = [
  'CR - IN - 3digits',
  'CR - IN - 4digits'
];

// turn an object style name (a string) into an object style
function getObjectStyleByName(doc, objectStyleName) {
  for (var i = 0; i < doc.allObjectStyles.length; i++) {
    var os = doc.allObjectStyles[i];
    if (os.name == objectStyleName) {
      return os;
    }
  }

  return null;
}

function createProgressBarPanel(args) {
  // parse out the args
  var title = args.title || 'Progress';
  var maximumValue = args.maximumValue;
  var width = args.width;

  var progressBarPanel = new Window(
    'window',
    title,
    undefined,
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

function main() {
  var doc = app.activeDocument;
  var numPageItems = doc.allPageItems.length;

  // this takes *some* time, so make a progress bar
  var args = {
    maximumValue: numPageItems,
    width: 400,
    title: 'Clearing object overrides. . . '
  };
  progressBarPanel = createProgressBarPanel(args);
  progressBarPanel.show();

  var oStyles = [];
  for (var i = 0; i < O_STYLE_NAMES.length; i++) {
    oStyles.push(getObjectStyleByName(doc, O_STYLE_NAMES[i]));
  }

  // loop over all the page items in this doc.
  for (var i = 0; i < numPageItems; i++) {
    // update the progress bar
    progressBarPanel.progressBar.value = i;
    var pageItem = doc.allPageItems[i];

    // for each page item, loop over the object styles we care about
    for (var j = 0; j < oStyles.length; j++) {
      var oStyle = oStyles[j];
      if (pageItem.appliedObjectStyle == oStyle) {
        // re-apply the object style, clearing overrides.
        pageItem.applyObjectStyle(oStyle, true);

        // each page item can only have one applied object style, so no need
        // to check if this matches the rest.
        break;
      }
    }
  }

  // clean up after ourselves
  progressBarPanel.close();
  alert('done');
}

// run it
main();
