// var label = 'captions';
var doc = app.activeDocument;

var curr = doc.pages[0].textFrames[0];

for (var i = 1; i < doc.pages.length; i++) {
  var page = doc.pages[i];
  var next = page.textFrames[0];

  if (next) {
    curr.nextTextFrame = next;
  }

  // for (var j = 1; j >= 0; j--) {
  //   var next = page.textFrames[j];
  //   curr.nextTextFrame = next;
  //   curr = next;
  // }
}

alert('done');
