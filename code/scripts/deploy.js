// global Scripting.FileSystemObject
var fso = WScript.CreateObject("Scripting.FileSystemObject");
// global shell.application
var shellApp =  WScript.CreateObject("shell.application");
// WScript.Shell
var shell = WScript.CreateObject("WScript.Shell");

//------------------------------------------------------------------------------
function getArgs(options, errors) {
  var args = {};
  for (var i in options) {
    if (options.hasOwnProperty(i)) {
      var option = options[i];
      try {
        if (!WScript.Arguments.Named.Exists(i)) {
          throw new Error("missing");
        }
        switch(option.type) {
          case 'file':
            args[i] = fso.GetFile(fso.GetAbsolutePathName(WScript.Arguments.Named.Item(i)));
            break;
          case 'folder':
            args[i] = fso.GetFolder(fso.GetAbsolutePathName(WScript.Arguments.Named.Item(i)));
            break;
          default:
            args[i] = WScript.Arguments.Named.Item(i);
            break;
        }
      }
      catch(e) {
        errors.push("Argument '" + i + "' (" + option.name + '): ' + e.description + '.');
      }
    }
  }
  return args;
}

//------------------------------------------------------------------------------
function createFolderR(folderName) {
  if (!fso.FolderExists(folderName)) {
    createFolderR(fso.GetParentFolderName(folderName));
    fso.CreateFolder(folderName);
  }
  return fso.GetFolder(folderName);
}

//------------------------------------------------------------------------------
function openInputFile(fileName) {
  var file = WScript.CreateObject('Msxml2.DOMDocument.6.0');
  file.async = false;
  if (!file.load(fileName)) {
    return null;
  }
  file.setProperty('SelectionLanguage', 'XPath');
  file.setProperty('SelectionNamespaces', 'xmlns:a="http://schemas.microsoft.com/developer/msbuild/2003"');
  return file;
}

function doCopyOperation(srcName, dstName) {
WScript.Echo('Copy ' + srcName + ' => ' + dstName);
  var src = fso.FolderExists(srcName) ? fso.GetFolder(srcName) : fso.FileExists(srcName) ? fso.GetFile(srcName) : null;
  var dst = fso.FolderExists(dstName) ? fso.GetFolder(dstName) : createFolderR(dstName);
  if (!src || !dst) {
    throw new Error('Missing source or destination');
  }
  if (fso.FolderExists(srcName)) {
    // src is a folder
    src.Copy(dst.Path);
  }
  else {
    // src is a file
    src.Copy(dst.Path + '\\');
  }
}

//------------------------------------------------------------------------------
var errors = [];
var args = getArgs({
  // build folder: major source of files
  'b': {
    'name': 'build folder',
    'type': 'folder'
  },
  // deploy folder: root for output
  'd': {
    'name': 'deploy folder',
    'type': 'folder'
  },
  // config file
  'c': {
    'name': 'config file',
    'type': 'file'
  },
  's': {
    'name': 'settings',
    'type': 'file'
  }
}, errors);

if (errors.length) {
  WScript.Echo('Errors:');
  for (var i = 0; i < errors.length; ++i) {
    WScript.Echo(errors[i]);
  }
  WScript.Quit(-1);
}

var projectName = fso.GetBaseName(args.c.Path);
WScript.Echo('Projectname: "' + projectName + '"');

var srcFolder = fso.GetFolder(shell.CurrentDirectory);

//------------------------------------------------------------------------------
// open config file
var inputFile = openInputFile(args.c.Path);

// open settings file
var settingsFile = openInputFile(args.s.Path);

//------------------------------------------------------------------------------
// get version from "/deploy/version"
var version = (function getVersion() {
  function parseVersion(s) {
    var m = parseVersion.exp.exec(s);
    return (m) ?
      {
        major: m[1],
        minor: m[2],
        patch: m[3],
        pre:   m[5] || '',
        meta:  m[7] || ''
      } :
      null;
  }
  parseVersion.exp = /^([0-9]+)\.([0-9]+)\.([0-9]+)(-([0-9A-Za-z-\.]+))?(\+([0-9A-Za-z-\.]+))?$/;

  var _version = parseVersion('0.1.0-pre');
  var nodeList = settingsFile.selectNodes('/a:Project/a:PropertyGroup[@Label="UserMacros"]/a:' + projectName + 'Version');
  if (nodeList.length) {
    var node = nodeList.nextNode();
    if(node) {
      _version = parseVersion(node.text);
    }
  }

  _version.toString = function() {
    var s = "" + this.major + '.' + this.minor + '.' + this.patch;
    s += (this.pre) ? '-' + this.pre : '';
    s += (this.meta) ? '+' + this.meta : '';
    return s;
  };

  return _version;
})();
WScript.Echo('Version: ' + version);

//------------------------------------------------------------------------------
// parse "/deploy/include/config" nodes
var configs = (function getConfigs() {
  var configNameMap = {};

  var nodeList = inputFile.selectNodes('/deploy/include/config');
  if (nodeList.length) {
    var node = nodeList.nextNode();
    while(node) {
      configNameMap[node.getAttribute('config') + '_' + node.getAttribute('platform')] = true;
      node = nodeList.nextNode();
    }
  }

  var configNames = [];
  for (var i in configNameMap) {
    configNames.push(i);
  }

  return configNames;
})();
for (var i = 0; i < configs.length; i++) {
  WScript.Echo('Adding ' + configs[i]);
}

//------------------------------------------------------------------------------
// copy files
var myBuildFolder = args.d.Path + '\\' + projectName + '-' + version;
if (fso.FolderExists(myBuildFolder)) {
  myBuildFolder = fso.GetFolder(myBuildFolder);
  fso.DeleteFolder(myBuildFolder.Path + '\\*.*');
}
else {
  myBuildFolder = fso.CreateFolder(myBuildFolder);
}
WScript.Echo('Copying files to ' + myBuildFolder.Path);

(function(){
  var nodeList = inputFile.selectNodes('/deploy/copy/item');
  if (nodeList.length) {
    var node = nodeList.nextNode();
    while(node) {
      var src = node.getAttribute('src');
      var dst = node.getAttribute('dst');
      if (src && dst) {
        var srcName = fso.GetAbsolutePathName(srcFolder.Path + '\\' + src);
        var dstName = fso.GetAbsolutePathName(myBuildFolder.Path + '\\' + dst);
        doCopyOperation(srcName, dstName);
      }
      node = nodeList.nextNode();
    }
  }
})();

//------------------------------------------------------------------------------
// create zip
var outputFile = args.d.Path + '\\' + projectName + '-' + version + '.zip';
(function(){
  var file = fso.CreateTextFile(outputFile, true)
  file.write("PK" + String.fromCharCode(5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
  file.close();
})();

//------------------------------------------------------------------------------
// copy to zip
WScript.StdOut.Write('Copying files to zip ' + outputFile);
(function(){
  var src = shellApp.NameSpace(myBuildFolder.Path).Items();
  var dst = shellApp.NameSpace(outputFile);
  dst.CopyHere(src, 2048 | 1024 | 512 | 16 | 4);
  var srcCount = src.count;
  while(dst.Items().count < srcCount) {
    WScript.Sleep(500);
    WScript.StdOut.Write('.');
  }
})();

WScript.Echo("\nDone.");
