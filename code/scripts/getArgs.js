//------------------------------------------------------------------------------
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

getArgs._default = function(aValue) {
  return aValue;
};



