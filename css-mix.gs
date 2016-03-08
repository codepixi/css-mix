var body;
function onOpen(e) {

  body = DocumentApp.getActiveDocument().getBody();
  //Browser.msgBox("body");
  log('coucou');
  //var test = body.getChild(0);
  //log(test.getType());

  // https://developers.google.com/apps-script/reference/document/paragraph#setattributesattributes
  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;

  var styleHeading1 = Object.create(style);
  styleHeading1[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ff0000';
  var styleHeading2 = Object.create(style);
  styleHeading2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#00ff00';
  

  var range = null;
  while(range = body.findElement(DocumentApp.ElementType.PARAGRAPH, range))
  {
    //range.getElement().asParagraph().setBold(true);
    paragraph = range.getElement().asParagraph();
    Logger.log(paragraph.getHeading());
    // https://developers.google.com/apps-script/reference/document/paragraph-heading
    switch(paragraph.getHeading())
    {
      case DocumentApp.ParagraphHeading.HEADING1: 
        paragraph.setAttributes(styleHeading1);
      break;
      case DocumentApp.ParagraphHeading.HEADING2: 
        paragraph.setAttributes(styleHeading2);
      break;
      default:break;
    }

    //log('while'); // create infinite loop
    //element = range.getElement();
    //log(element);
  }
}

function log(message)
{
    body.appendParagraph(message);
}

