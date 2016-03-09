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
function convertToGoogleStyle(userStyleList)
{
  var styleMap = {};
  styleMap['color'] = DocumentApp.Attribute.FOREGROUND_COLOR;

  googleStyleList = {};
  
  googleStyleList['h1'] = Object.create(style);
  googleStyleList['h1'][DocumentApp.Attribute.FOREGROUND_COLOR] = userStyleList['h1']['color'];
  googleStyleList['h2'] = Object.create(style);
  googleStyleList['h2'][DocumentApp.Attribute.FOREGROUND_COLOR] = userStyleList['h2']['color'];
  googleStyleList['h3'] = Object.create(style);
  googleStyleList['h3'][DocumentApp.Attribute.FOREGROUND_COLOR] = userStyleList['h3']['color'];
  return googleStyleList;
}

function applyStyleSheet(googleStyleMap)
{
  // https://developers.google.com/apps-script/reference/document/paragraph-heading
  var selectorMap = {};
  selectorMap[DocumentApp.ParagraphHeading.HEADING1] = 'h1';
  selectorMap[DocumentApp.ParagraphHeading.HEADING2] = 'h2';
  selectorMap[DocumentApp.ParagraphHeading.HEADING3] = 'h3';
  selectorMap[DocumentApp.ParagraphHeading.NORMAL] = 'p';
  
  var range = null;    
  if(!body) body = DocumentApp.getActiveDocument().getBody();
  var test = '';
  while(range = body.findElement(DocumentApp.ElementType.PARAGRAPH, range))
  {
    paragraph = range.getElement().asParagraph();
    Logger.log(paragraph.getHeading());
    style = googleStyleMap[selectorMap[paragraph.getHeading()]];
    if(style) {paragraph.setAttributes(style);}
  }
  log(test);
}

var cssForm;
function prepareStyleEntry() 
{
  //https://developers.google.com/apps-script/guides/html/reference/run#methods
  cssForm = HtmlService.createHtmlOutputFromFile('CssForm')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Your CSS StyleSheet')
      .setWidth(300);
  DocumentApp.getUi().showSidebar(cssForm);
  //onclick="google.script.host.close()"
}

//https://developers.google.com/apps-script/troubleshooting#common_errors
function askStyleSheet()
{
  //getOAuthToken();
  //https://developers.google.com/apps-script/reference/base/prompt-response
  var styleSheet = "h1 {color:#817DF5;} h2 {color:#F7B2CE;}";
  //https://code.google.com/p/google-apps-script-issues/issues/detail?id=677
  //var styleSheet = DocumentApp.getUi().prompt('Write your stylesheet within one line, no space around the : symbol').getResponseText();
  return styleSheet;
}

function onOpen(e) 
{
  body = DocumentApp.getActiveDocument().getBody();
  //log(e.authMode);
  prepareStyleEntry();
  //styleSheet = askStyleSheet();
  //processStyleSheet(styleSheet);
}

function processStyleSheet(styleSheet)
{
  //log("processStyleSheet");
  //log(styleSheet);
  styleSheet = styleSheet.replace(/(?:\r\n|\r|\n)/g, '');
  userStyleMap = parseStyleSheet(styleSheet);
  //log(JSON.stringify(userStyleMap));
  googleStyleMap = convertToGoogleStyle(userStyleMap);
  applyStyleSheet(googleStyleMap);
}

function log(message)
{
    if(!body) body = DocumentApp.getActiveDocument().getBody();
    body.appendParagraph(message);
}
