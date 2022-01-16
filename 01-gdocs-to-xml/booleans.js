function isParagraph(element) {
  return element && element.getType() === DocumentApp.ElementType.PARAGRAPH;
}

function isListItem(element) {
  return element && element.getType() === DocumentApp.ElementType.LIST_ITEM;
}

function isText(element) {
  return element && element.getType() === DocumentApp.ElementType.TEXT;
}

function isPageBreak(element) {
  return element && element.getType() === DocumentApp.ElementType.PAGE_BREAK;
}

function isEmpty(element) {
  return isText(element) && /^\s*$/gi.test(element.getText());
}
