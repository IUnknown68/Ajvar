var fso = WScript.CreateObject("Scripting.FileSystemObject");

//------------------------------------------------------------------------------
function createFolderR(folderName) {
  if (!fso.FolderExists(folderName)) {
    createFolderR(fso.GetParentFolderName(folderName));
    fso.CreateFolder(folderName);
  }
  return fso.GetFolder(folderName);
}

