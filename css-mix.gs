var body;

function parseStyleSheet(styleSheet)
{
  var userDeclarationList = styleSheet.split('}');
  var declarationSyntax = new RegExp("([a-zA-Z0-9.#-]*) *{( *[^ ]* *)}");
  var userStyleMap = {};
  
  for(position in userDeclarationList)
  {
    userDeclaration = userDeclarationList[position] + '}';
    if(userDeclaration == '}') break;
    //log('userDeclaration ' + userDeclaration);
    splicedDeclaration =  declarationSyntax.exec(userDeclaration);
    selector = splicedDeclaration[1];
    //log("selector " + selector);
    //log("statement list " + splicedDeclaration[2]);
    userStyleMap[selector] = {};
    statements = splicedDeclaration[2].split(';');
    for(position in statements)
    {
      //log('-' + statements[position] + '-');
      if(statements[position])
      {
        couple = statements[position].split(':');
        userStyleMap[selector] = {};
        userStyleMap[selector][couple[0]] = couple[1];
      }
    }
  }
  return userStyleMap;
}

var style = {};
//style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.RIGHT;
style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
//style[DocumentApp.Attribute.FONT_SIZE] = 18;
//style[DocumentApp.Attribute.BOLD] = true;

// https://developers.google.com/apps-script/reference/document/paragraph#setattributesattributes
function convertToGoogleStyle(userStyleMap)
{
  googleStyleMap = {};
  googleStyleMap['h1'] = Object.create(style);
  googleStyleMap['h1'][DocumentApp.Attribute.FOREGROUND_COLOR] = userStyleMap['h1']['color'];
  googleStyleMap['h2'] = Object.create(style);
  googleStyleMap['h2'][DocumentApp.Attribute.FOREGROUND_COLOR] = userStyleMap['h2']['color'];
  return googleStyleMap;
}

function onOpen(e) {
  body = DocumentApp.getActiveDocument().getBody();
  var styleSheet = "h1 {color:#FAC941;} h2 {color:#8F6D0E;}";
  userStyleMap = parseStyleSheet(styleSheet);
  log(JSON.stringify(userStyleMap));
  googleStyleMap = convertToGoogleStyle(userStyleMap);

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
        paragraph.setAttributes(googleStyleMap['h1']);
      break;
      case DocumentApp.ParagraphHeading.HEADING2: 
        paragraph.setAttributes(googleStyleMap['h2']);
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

