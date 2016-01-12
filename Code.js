/** 
 * URL Shortener: Shorten long URLs using Google's URL Shortener service
 * Creator: Allen Guo
 * License: MIT
 * Website: https://github.com/guoguo12/url-shortener
 */

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
                     .addItem('Get document link', 'showDocDialog')
                     .addItem('Shorten selected URL', 'showSelectedDialog')
                     .addItem('Shorten and replace selected URL', 'shortenAndReplace')
                     .addToUi();
}

function onInstall() {
  onOpen();
}

function showDocDialog() {
  var dialog = HtmlService
    .createHtmlOutputFromFile('DocDialog')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  dialog.setHeight(88);
  dialog.setWidth(375);
  DocumentApp.getUi().showModalDialog(dialog, 'Get document link');
}

function showSelectedDialog() {
  var dialog = HtmlService
    .createHtmlOutputFromFile('SelectedDialog')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  dialog.setHeight(206);
  dialog.setWidth(479);
  DocumentApp.getUi().showModalDialog(dialog, 'Shorten selected URL');
}

function shortenAndReplace() {
  var urls = shortenSelUrl();
  if (urls) {
    replaceSelectedInSelection(urls[0], urls[1]);
  }
}

/**
 * Returns the shortened document URL. Fetches the shortened URL if it is not
 * already saved in the document properties.
 */
function shortenDocUrl() {
  var props = PropertiesService.getDocumentProperties();
  var savedUrl = props.getProperty('short-url');
  if (!savedUrl) {
    var docUrl = DocumentApp.getActiveDocument().getUrl();
    savedUrl = shorten(docUrl);
    props.setProperty('short-url', savedUrl);
  }
  return savedUrl;
}

/**
 * Returns a two-element array containing the selected URL and its
 * shortened version.
 */
function shortenSelUrl() {
  // Extract URL from selection
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    showError('Nothing is selected.');
    return;
  }
  var elements = selection.getRangeElements();
  if (elements.length < 1) {
    showError('No element selected.');
    return;      
  }
  if (elements.length > 1) {
    showError('Too many elements are selected.');
    return;      
  }
  var rangeElem = elements[0];
  var elem = rangeElem.getElement();
  try {
    elem.editAsText();  
  } catch (e) {
    showError('Selection is not a valid URL.');
    return;
  }
  var selUrl = trim(getTextFromElement(elem, rangeElem));
  if (!selUrl || selUrl.indexOf(' ') != -1 || selUrl.indexOf('\\') != -1 || selUrl.indexOf('.') == -1) {
    showError('Selection does not appear to be a valid URL.');
    return;   
  }
  // Already-shortened links cannot be shortened again
  if (selUrl.indexOf('goo.gl') != -1) {
    showError('This URL cannot be shortened.');
    return;
  }
  // Fetch shortened URL, throwing errors to failure handler above
  var shortUrl = shorten(selUrl);
  return [selUrl, shortUrl]
}

/** 
 * Returns the shortened version of the given URL.
 */
function shorten(inputUrl) {
  // TODO: Add additional tests for clearly invalid URLs.
  try {
    var data = UrlShortener.Url.insert({longUrl: inputUrl});
  } catch (e) {
    throw 'URL Shortener service request failed (reason: "' + e + '")';
  }
  if (data.id) {
    return data.id;
  } else {
    throw 'URL Shortener service request failed';
  }
}

/**
 * Replaces all instances of the given word in all elements that are partially
 * or fully selected.
 */
function replaceSelectedInSelection(original, shortened) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    return;
  }
  var elements = selection.getRangeElements();
  for (var i = 0; i < elements.length; i++) {
    try {
      var text = elements[i].getElement().editAsText();
      text.replaceText(escape(original), shortened);
    } catch (e) {
      // Element is not text, so continue
    }
  }
}

function showWarning(message) {
  var ui = DocumentApp.getUi();
  var response = ui.alert('Warning', message, ui.ButtonSet.YES_NO);
  return response == ui.Button.YES;
}

function showError(message) {
  var ui = DocumentApp.getUi();
  ui.alert('Error', message, ui.ButtonSet.OK); 
}

function getTextFromElement(elem, rangeElem) {
  var start = rangeElem.getStartOffset();
  var end = rangeElem.getEndOffsetInclusive() + 1;
  return elem.editAsText().getText().substring(start, end);
}

/**
 * Unused.
 */
function isLikelyValidUrl(str) {
  var pattern = /(http|ftp|https):\/\/[\w-]+(\.[\w-]+)+([\w.,@?^=%&amp;:\/~+#-]*[\w@?^=%&amp;\/~+#-])?/;
  var match = str.match(pattern);
  return match != null && str == match[0];
}

/**
 * Removes surrounding whitespace.
 */
function trim(str) {
 return str.replace(/^\s+|\s+$/g, '');
}

/**
 * Escapes a non-regex string for exact matching using regex.
 */
function escape(str) {
  return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}
