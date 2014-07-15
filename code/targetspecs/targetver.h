#pragma once

// Including SDKDDKVer.h defines the highest available Windows platform.

// If you wish to build your application for a previous Windows platform, include WinSDKVer.h and
// set the _WIN32_WINNT macro to the platform you wish to support before including SDKDDKVer.h.


#include <WinSDKVer.h>

// minimum requirements:
// _WIN32_WINNT_VISTA - Vista
// _WIN32_IE_IE90 - IE9

#ifndef _WIN32_WINNT
# define _WIN32_WINNT  0x0600
#endif

#ifndef _WIN32_IE
# define _WIN32_IE     0x0900
#endif

#include <SDKDDKVer.h>