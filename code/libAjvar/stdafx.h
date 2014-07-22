// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include <targetver.h>

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS  // some CString constructors will be explicit
#define _ATL_IGNORE_NAMED_ARGS

// ATL / WTL
#include <atlbase.h>
#include <atlstr.h>
#include <atlcom.h>
#include <atlsafe.h>

// Windows stuff
#include <Exdisp.h>
#include <dispex.h>

// std::
#include <string>
#include <map>
#include <set>
#include <unordered_map>
#include <algorithm>
#include <functional>

// gtest
#include "gtest/gtest.h"
#include "gmock/gmock.h"

