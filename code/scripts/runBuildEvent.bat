@echo off
set file=%1
shift
set major=%1
shift
set minor=%1
shift
set patch=%1
shift
set pre=%1
shift
set meta=%1
shift
set project=%1
shift
set event=%1
shift
set platform=%1
shift
set configuration=%1
shift
set solution=%1

if exist %file% (
  cscript %file% %major% %minor% %patch% %pre% %meta% %project% %event% %platform% %configuration% %solution%
)