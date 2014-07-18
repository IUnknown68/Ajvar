// global Scripting.FileSystemObject
var fso = WScript.CreateObject("Scripting.FileSystemObject");
// global shell.application
var shellApp =  WScript.CreateObject("shell.application");
// WScript.Shell
var shell = WScript.CreateObject("WScript.Shell");




if (!WScript.Arguments.Unnamed.Length) {
  WScript.Echo('Error: Missing argument 0: hookname');
  WScript.Quit(-1);
}

var hookName = WScript.Arguments

WScript.Quit(-1);



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
var errors = [];
var args = getArgs({
  's': {
    'name': 'source folder',
    'type': 'folder'
  }
}, errors);

if (errors.length) {
  WScript.Echo('Errors:');
  for (var i = 0; i < errors.length; ++i) {
    WScript.Echo(errors[i]);
  }
  WScript.Quit(-1);
}

//------------------------------------------------------------------------------
// create zip
var outputFile = args.s.Path + '.zip';

(function(){
  var file = fso.CreateTextFile(outputFile, true);
  file.write("PK" + String.fromCharCode(5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
  file.close();
})();

//------------------------------------------------------------------------------
// copy to zip
WScript.StdOut.Write('Copying files to zip ' + outputFile);
(function(){
  var src = shellApp.NameSpace(args.s.Path).Items();
  var dst = shellApp.NameSpace(outputFile);
  dst.CopyHere(src, 2048 | 1024 | 512 | 16 | 4);
  var srcCount = src.count;
  while(dst.Items().count < srcCount) {
    WScript.Sleep(500);
    WScript.StdOut.Write('.');
  }
})();

WScript.Echo("\nDone.");
