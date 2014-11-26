About
=====

Ajvar is a library containing helpers for dealing with ActiveScript-related
interfaces. It is focused on `IDispatch`, `IDispatchEx` and the webbrowser
control (`IWebBrowser`).

Requirements
============
- Visual Studio 2012 or higher (untested, 2012 used for development).
- [Gtest](http://code.google.com/p/googletest/) and [gmock](http://code.google.com/p/googlemock/).
- For building the documentation: [doxygen](http://www.stack.nl/~dimitri/doxygen/) 1.8.6 or higher.

Installation
============
- Clone this repository.
- Copy `code\settings.props` to `code\settings.user.props`.
- Open the project in VS.
- Open the "Property Manager" in VS and find the file `code\settings.user`
  (the file you just copied).
- Adjust the paths in the section "User Macros" according to your project layout.

Generally we use foldernames for build folders (and by this for libs) following this format:

`$(Configuration)_$(PlatformName)`

So e.g. for the 32bit debug version, output will be generated in
`Debug_Win32`, and libraries are expected in a some folder named equally.
The `$(FullConfiguration)` macro can be used to refer to this combination.

The Ajvar solution file contains the following projects:

1) libAjvar
   The actual ajvar libraries.

2) libAjvarTests
   Gtest-based unit tests. Builds and runs them.
   Requires: gtest, gmock (see below).

3) Documentation
   Yes, that's right - the documentation. Uses doxygen to build the html based documentation.
   Requires doxygen.

4) Deploy
   A helper project to put everything together: libraries, documentation, all neatly placed in a folder named 'Ajvar-<version>', ready to zip.
   Note: I didn't include code to actually create the zip, because I couldn't find a good way to do this with windows' built-in functions and I don't want to require any additional external tools.
   There is some code though to create the zip via the windows shell, but it turned out, there is no way to detect when zipping finished - and when you quit the calling process too early, your zip will be incomplete.
   I guess it is acceptable, if you want a zip, to manually rightclick the folder and select "Add to zip" from the context menu.

Building
========
##### Libraries / tests
Do a "Batch build" in VS and check `libAjvarTests` for in each configuration you
need. The dependencies will also build `libAjvar`.

##### Documentation
Build the "Documentation" project. It depends on `libAjvarTests`, so
`libAjvar` and `libAjvarTests` will be built in the selected configuration.
You need [doxygen](http://www.stack.nl/~dimitri/doxygen/) for this.

##### Deploy
You can build all "Deploy" targets as a batch. This will generate
`deploy/Ajvar-<VERSION>` containing the libraries, include files and the documentation.

```
deploy/Ajvar-<VERSION>
  README.md     This file
  Ajvar.html    Documentation
  docs/         Documentation files
  include/      Include files
    Browser/
    Dispatch/
    ...
  lib/          Libraries
    Debug_Win32/
    Release_Win32/
    Debug_x64/
    Release_x64/
  tests/        Tests
    Debug_Win32/
    Release_Win32/
    Debug_x64/
    Release_x64/
```

All builds of `libAjvarTests` will also run the tests.

Buildsystem
-----------
The buildsystem is based on `$(SolutionDir)events.props`. This file handles the
build events `PreBuild`, `PreLink` and `PostBuild`. For each of them it runs,
if present, a javascript file `$(ProjectDir)events\<EVENTNAME>.wsf` performing
tasks like deploying.

API
===

Types:
------
- `Ajvar::ComVariant`
- `Ajvar::ComBSTR`
- `Ajvar::Time::Win32::Time`
- `Ajvar::Time::Unix::Time`
- `Ajvar::Time::JScript::Time`

Constants:
----------
- `Ajvar::Time::Timespan`
- `Ajvar::Time::Minute`
- `Ajvar::Time::Hour`
- `Ajvar::Time::Day`
- `Ajvar::Time::Week`
- `Ajvar::Time::ShortMonth`
- `Ajvar::Time::LongMonth`
- `Ajvar::Time::Year`

Functions:
----------
- `Ajvar::Time::Win32::from(Unix::Time)`
- `Ajvar::Time::Win32::from(JScript::Time)`
- `Ajvar::Time::Win32::now()`
- `Ajvar::Time::Win32::from(Win32::Time)`
- `Ajvar::Time::Win32::from(JScript::Time)`
- `Ajvar::Time::Win32::now()`
- `Ajvar::Time::Win32::from(Unix::Time)`
- `Ajvar::Time::Win32::from(Win32::Time)`
- `Ajvar::Time::Win32::now()`
- `Ajvar::Dispatch::GetScriptDispatch(IWebBrowser2 *, TInterface **)`
- `Ajvar::Dispatch::Ex::EachDispEx(HRESULT, const BSTR &, const DISPID, const VARIANT  &)`
- `Ajvar::Dispatch::Ex::forEach(IDispatchEx *, EachDispEx, DWORD)`

Classes:
--------
- `Ajvar::Dispatch::Object`
- `Ajvar::Dispatch::RefVariant`
- `Ajvar::Dispatch::Ex::ObjectGet`
- `Ajvar::Dispatch::Ex::ObjectCreate`
- `Ajvar::Dispatch::Ex::Object`
- `Ajvar::Dispatch::Ex::Properties`


Templates:
----------
- `Ajvar::Dispatch::_Connector`
