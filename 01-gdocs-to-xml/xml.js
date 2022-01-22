const YEAR_STYLE = DocumentApp.ParagraphHeading.HEADING1;

function createXml() {

  let root = XmlService.createElement('document');

  const gdoc = DocumentApp.getActiveDocument();
  const body = gdoc.getBody();
  const numChildren = body.getNumChildren();
  let listIndex = 1;

  for (let i = 0; i < numChildren; i++) {

    let gdchild = body.getChild(i);

    if (isListItem(gdchild)) {
      const xmlentry = createEntryFromListItem(gdchild, listIndex);
      root.addContent(xmlentry);
      listIndex++;
    } else if (isParagraph(gdchild) && gdchild.getHeading() == YEAR_STYLE) {
      let misc = XmlService.createElement('h1')
        .setText(gdchild.getText())
      root.addContent(misc);
    }
  }

  let document = XmlService.createDocument(root);
  let xml = XmlService.getPrettyFormat().format(document);

  return xml;
}

function createEntryFromListItem(listItem, listIndex) {

  let xmlentry = XmlService.createElement('entry');
  let xmlnumber = XmlService.createElement('number')
    .setText(listIndex);
  xmlentry.addContent(xmlnumber);

  return xmlentry;
}

function xmlTest() {

  const xmlProlog = '<?xml version="1.0" encoding="UTF-8"?>\n';
  const xmlFragment = `${xmlProlog}<entry>some text <italic>with</italic> tags</entry>`;
  var fragment = XmlService.parse(xmlFragment);
  var fragmentRoot = fragment.detachRootElement();
  
  let root = XmlService.createElement('document');
  root.addContent(fragmentRoot);
  let document = XmlService.createDocument(root);
  let xml = XmlService.getPrettyFormat().format(document);

  Logger.log(xml);
}

function xmlTest2() {
  const fragment = '<entry>some text <italic>with</italic> tags</entry>';
  const el = parseFragment(fragment);

  const root = XmlService.createElement('document');
  root.addContent(el);
  const document = XmlService.createDocument(root);
  const xml = XmlService.getPrettyFormat().format(document);
  Logger.log(xml);
}

function parseFragment(fragment) {
  const xmlProlog = '<?xml version="1.0" encoding="UTF-8"?>\n';
  const xmlDoc = `${xmlProlog}${fragment}`;
  return XmlService.parse(xmlDoc).detachRootElement();
}
