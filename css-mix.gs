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

function applyStyleSheet(googleStyleMap)
{
  var range = null;    
  if(!body) body = DocumentApp.getActiveDocument().getBody();
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
  }
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
}

function processStyleSheet()
{
  //log("processStyleSheet");
  styleSheet = askStyleSheet();
  //log(styleSheet);
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

function getAuthentificationToken() 
{
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}