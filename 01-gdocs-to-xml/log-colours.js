function logColours() {

  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  const numChildren = body.getNumChildren();
  const bgColours = [];
  const fgColours = [];

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const numGrandchildren = child.getNumChildren();
    for (let j = 0; j < numGrandchildren; j++) {
      const grandchild = child.getChild(j);
      if (grandchild.getType() == 'TEXT') {
        const indices = grandchild.getTextAttributeIndices();
        for (let index of indices) {
          let atts = grandchild.getAttributes(index);
          if (atts.FOREGROUND_COLOR && !fgColours.includes(atts.FOREGROUND_COLOR)) {
            fgColours.push(atts.FOREGROUND_COLOR);
          }
          if (atts.BACKGROUND_COLOR && !bgColours.includes(atts.BACKGROUND_COLOR)) {
            bgColours.push(atts.BACKGROUND_COLOR);
          }
        }
      }
    }
  }
  Logger.log(JSON.stringify({
    fgColours,
    bgColours,
  }, null, '  '));
}
