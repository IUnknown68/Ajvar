var settings = getArgs({
  'settings': {
    name: 'settings file',
    type: getArgs.settings
  }
}).settings;

settings.version = {
  major: settings.get('Version_Major'),
  minor: settings.get('Version_Minor'),
  patch: settings.get('Version_Patch'),
  pre: settings.get('Version_Pre'),
  meta: settings.get('Version_Meta'),
  toString: function() {
    var s = '' + this.major + '.' + this.minor + '.' + this.patch;
    if (this.pre) {
      s += '-' + this.pre;
    }
    if (this.meta) {
      s += '+' + this.meta;
    }
    return s;
  }
};

