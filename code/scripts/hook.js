//------------------------------------------------------------------------------
(function(){
  var errors = [];
  g_version = getArgs({
    'major': {
      name: 'major version',
      type: getArgs.int10
    },
    'minor': {
      name: 'minor version',
      type: getArgs.int10
    },
    'patch': {
      name: 'patch version',
      type: getArgs.int10
    },
    'pre': {
      name: 'pre version'
    },
    'meta': {
      name: 'meta version'
    }
  }, errors);
  g_version.toString = function() {
    var s = '' + this.major + '.' + this.minor + '.' + this.patch;
    if (this.pre) {
      s += '-' + this.pre;
    }
    if (this.meta) {
      s += '+' + this.meta;
    }
     return s;
  };

  g_args = getArgs({
    'event': {
      name: 'event name'
    },
    'project': {
      name: 'project name'
    },
    'platform': {
      name: 'platform name'
    },
    'configuration': {
      name: 'configuration name'
    },
    'solution': {
      name: 'solution name'
    }
  }, errors);

  if (errors.length) {
    WScript.Echo('Errors:');
    for (var i = 0; i < errors.length; ++i) {
      WScript.Echo(errors[i]);
    }
    WScript.Quit(-1);
  }
})();

