//------------------------------------------------------------------------------
/// @function getArgs
/// @brief Gets the command line arguments
function getArgs(options, errors) {
  var args = {};
  for (var i in options) {
    if (options.hasOwnProperty(i)) {
      var option = options[i];
      try {
        var argName = WScript.Arguments.Named.Exists(i) ?
          i :
            (option.alias && WScript.Arguments.Named.Exists(option.alias)) ?
            option.alias :
            null;
        if (!argName) {
          throw new Error("is missing");
        }
        args[i] = (option.type || getArgs._default)(WScript.Arguments.Named.Item(argName));
      }
      catch(e) {
        var s = "Argument '" + i + "' (" + option.name + ') ' + e.description + '.';
        if (!errors) {
          throw new Error(s);
        }
        errors.push(s);
      }
    }
  }
  return args;
}

getArgs.file = function(aValue) {
  return fso.GetFile(fso.GetAbsolutePathName(aValue));
};

getArgs.folder = function(aValue) {
  return fso.GetFolder(fso.GetAbsolutePathName(aValue));
};

getArgs.float = function(aValue) {
  aValue = parseFloat(aValue);
  if (isNaN(aValue)) {
    throw new Error('is not a number');
  }
  return aValue;
};

getArgs.intN = function(aValue, aBase) {
  aValue = parseInt(aValue, aBase);
  if (isNaN(aValue)) {
    throw new Error('is not a number');
  }
  return aValue;
};

getArgs.int10 = function(aValue) {
  return getArgs.intN(aValue, 10);
};

getArgs.int16 = function(aValue) {
  return getArgs.intN(aValue, 16);
};

getArgs.settings = function(aValue) {
  var file = fso.GetFile(fso.GetAbsolutePathName(aValue));

  var settingsFile = WScript.CreateObject('Msxml2.DOMDocument.6.0');
  settingsFile.async = false;
  if (!settingsFile.load(file.Path)) {
    throw new Error('failed loading');
  }
  settingsFile.setProperty('SelectionLanguage', 'XPath');
  settingsFile.setProperty('SelectionNamespaces', 'xmlns:a="http://schemas.microsoft.com/developer/msbuild/2003"');
  return {
    get: function(aName) {
      var nodeList = settingsFile.selectNodes('/a:Project/a:PropertyGroup[@Label="UserMacros"]/a:' + aName);
      var node;
      if (node = nodeList.nextNode()) {
        return node.text;
      }
      return undefined;
    }
  };
};

getArgs._default = function(aValue) {
  return aValue;
};
