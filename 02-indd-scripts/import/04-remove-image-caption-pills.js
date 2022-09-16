var doc = app.activeDocument;
// var selection = app.selection[0];

// alert(selection.constructor.name);

var textFramesToDelete = [];

for (var j = 0; j < doc.textFrames.length; j++) {
  var parentTextFrame = doc.textFrames[j];
  for (var i = 0; i < parentTextFrame.textFrames.length; i++) {
    var textFrame = parentTextFrame.textFrames[i];
    if (
      textFrame.appliedObjectStyle
      && textFrame.appliedObjectStyle.name == 'CR - IN - inline image'
      && !/\d/g.test(textFrame.contents)
    ) {
      textFramesToDelete.push(textFrame);
    }
  }
}

for (var i = 0; i < textFramesToDelete.length; i++) {
  var textFrame = textFramesToDelete[i];
  textFrame.remove();
}
