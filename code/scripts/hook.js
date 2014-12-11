/***
  A hook requires the following command line arguments:
  - everything from settings.js
  - 'event' : name of the event (PreBuild, PreLink, PostBuild)
  - 'project' : projectname
  - 'platform' : Win32 / x64
  - 'configuration' : e.g. Release, Debug
  - 'solution' : solutionname
  - 'deployfolder' : folder to deploy to, e.g. "deploy/sample-1.0.1-beta"
                     This will result in a zip file: "deploy/sample-1.0.1-beta.zip"
  - 'outfolder' : value of $(OutDir) (build folder)
  - 'intfolder' : value of $(IntDir) (intermediate folder)

Sample command line for your Build events (here: PreBuild):

if exist "$(SolutionDir)scripts\$(ProjectName)-PreBuild.wsf" (cscript "$(SolutionDir)scripts\$(ProjectName)-PreBuild.wsf" /event:PostBuild /project:"$(ProjectName)" /platform:"$(PlatformName)" /configuration:"$(Configuration)" /solution:"$(SolutionName)" /settings:"$(SolutionDir)settings.user.props" /deployfolder:"$(DeployVersionFolder)" /outfolder:"$(OutDir)" /intfolder:"$(IntDir)")

 */
//------------------------------------------------------------------------------
(function(){
  var errors = [];
  var args = getArgs({
    'event': {
      name: 'Event name'
    },
    'project': {
      name: 'Project name'
    },
    'platform': {
      name: 'Platform name'
    },
    'configuration': {
      name: 'Configuration name'
    },
    'solution': {
      name: 'Solution name'
    },
    'deployfolder': {
      name: 'Deploy folder',
      get: getArgs.folder,
      create: true
    },
    'buildfolder': {
      name: 'Build Folder',
      get: getArgs.folder
    },
    'outfolder': {
      name: 'Out Folder',
      get: getArgs.folder
    },
    'intfolder': {
      name: 'Intermediate folder',
      get: getArgs.folder
    }
  }, errors);

  if (errors.length) {
    WScript.Echo('Errors:');
    for (var i = 0; i < errors.length; ++i) {
      WScript.Echo(errors[i]);
    }
    WScript.Quit(-1);
  }
  // copy to global settings
  for (var i in args) {
    if (args.hasOwnProperty(i)) {
      settings[i] = args[i];
    }
  }
})();

