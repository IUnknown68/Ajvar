//------------------------------------------------------------------------------
(function(){
  var errors = [];
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
    },
    'settings': {
      name: 'settings file',
      type: getArgs.settings
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

