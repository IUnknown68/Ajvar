//------------------------------------------------------------------------------
function openXmlFile(fileName) {
  var file = WScript.CreateObject('Msxml2.DOMDocument.6.0');
  file.async = false;
  if (!file.load(fileName)) {
    return null;
  }
  file.setProperty('SelectionLanguage', 'XPath');
  file.setProperty('SelectionNamespaces', 'xmlns:a="http://schemas.microsoft.com/developer/msbuild/2003"');
  return file;
}
