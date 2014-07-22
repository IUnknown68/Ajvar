/**************************************************************************//**
@file
@brief     Main include file for using Ajvar
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
@mainpage  Ajvar DCOM helper library
@tableofcontents

@section About About
Ajvar is a library containing helpers for dealing with ActiveScript-related
things. It is focused on `IDispatch`, `IDispatchEx` and the webbrowser
control (`IWebBrowser`).

@section Types Types
- `Ajvar::ComVariant`
- `Ajvar::ComBSTR`
- `Ajvar::Time::Win32::Time`
- `Ajvar::Time::Unix::Time`
- `Ajvar::Time::JScript::Time`

@section Constants Constants
- `Ajvar::Time::Minute`
- `Ajvar::Time::Hour`
- `Ajvar::Time::Day`
- `Ajvar::Time::Week`
- `Ajvar::Time::ShortMonth`
- `Ajvar::Time::LongMonth`
- `Ajvar::Time::Year`

@section Functions Functions
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

- `Ajvar::Error::DErrFACILITY()`
- `Ajvar::Error::DErrWin32()`
- `Ajvar::Error::DErrSEVERITY()`
- `Ajvar::Error::DErrHRESULT()`

@section Classes Classes
- `Ajvar::Time::Timespan`

- `Ajvar::Dispatch::Object`
- `Ajvar::Dispatch::RefVariant`
- `Ajvar::Dispatch::Ex::ObjectGet`
- `Ajvar::Dispatch::Ex::ObjectCreate`
- `Ajvar::Dispatch::Ex::Object`
- `Ajvar::Dispatch::Ex::Properties`

- `Ajvar::Sync::CriticalSection`
- `Ajvar::Sync::Lock`

- `Ajvar::Debug::AjvarComObjectRootEx`

@section Templates Templates
- `Ajvar::Dispatch::_Connector`

@section Macros Macros
- `NO_COPY()`
- `AJ_ASSERT_()`
- `IF_FAILED_RET_HR()`
- `IF_FAILED_RET_ANY()`
- `IF_FAILED_BREAK()`
- `IF_NULL_RET()`
- `IS_VTERR()`

***************************************************************************/
#pragma once

/// @brief `NO_COPY` macro: Declares a private copy constructor and assignment
/// operator.
#ifndef NO_COPY
#define NO_COPY(_cls) \
  private: \
    _cls(const _cls &); \
    _cls & operator = (const _cls &);
#endif // ndef NO_COPY

/// @brief If `AJ_DEBUG_BREAK` is defined, the following error handling macros produce a
/// debug break.
#define AJ_DEBUG_BREAK

/// @brief `ASSERT` macro, asserts only when `AJ_DEBUG_BREAK` is defined.
#ifdef AJ_DEBUG_BREAK
# define AJ_ASSERT_ ATLASSERT
#else
# define AJ_ASSERT_
#endif

/// @brief Checks a `HRESULT` and returns if failed.
#define IF_FAILED_RET_HR(_hr) \
  do \
  { \
    HRESULT _hr__ = _hr; \
    if (FAILED(_hr__)) \
    { \
      ATLTRACE(_T("ASSERTION FAILED: 0x%08x (%s) in "), _hr__, DErrHRESULT(_hr__)); \
      ATLTRACE(__FILE__); \
      ATLTRACE(_T(" line %i\n"), __LINE__); \
      AJ_ASSERT_(false); \
      return _hr__; \
    } \
  } while(0);

/// @brief Checks a `HRESULT` and returns any value if failed.
#define IF_FAILED_RET_ANY(_hr, _ret) \
  do \
  { \
    HRESULT _hr__ = _hr; \
    if (FAILED(_hr__)) \
    { \
      ATLTRACE(_T("ASSERTION FAILED: 0x%08x (%s) in "), _hr__, DErrHRESULT(_hr__)); \
      ATLTRACE(__FILE__); \
      ATLTRACE(_T(" line %i\n"), __LINE__); \
      AJ_ASSERT_(false); \
      return _ret; \
    } \
  } while(0);

/// @brief Checks a `HRESULT` and breaks a loop if failed. `_hrRet` receives the error code.
#define IF_FAILED_BREAK(_hr, _hrRet) \
    _hrRet = _hr; \
    if (FAILED(_hrRet)) \
    { \
      ATLTRACE(_T("ASSERTION FAILED: 0x%08x (%s) in "), _hrRet, DErrHRESULT(_hrRet)); \
      ATLTRACE(__FILE__); \
      ATLTRACE(_T(" line %i\n"), __LINE__); \
      AJ_ASSERT_(false); \
      break; \
    }

/// @brief Checks a pointer value and returns `E_POINTER` if the value is `nullptr`.
#define IF_NULL_RET(_val) \
  if (nullptr == _val) { return E_POINTER; }

/// @brief Checks a `VARIANT` for being of type `VT_ERROR`.
#define IS_VTERR(_val) \
  (VT_ERROR == _val.vt)


//============================================================================
/// @brief Main namespace: ATL Extensions
/// @details  ATL Extensions inlcude a bunch of helpful stuff for dealing
/// with `IDispatch` and `IDispatchEx`. It uses ATL classes as baseclasses.
namespace Ajvar {

/// @brief Internally used smartVariant
typedef ATL::CComVariant ComVariant;

/// @brief Internally used smartBSTR
typedef ATL::CComBSTR ComBSTR;

} // namespace Ajvar

#include "Debug/ComObjectDebugging.h"

#include "Time/Time.h"

#include "Sync/CriticalSection.h"
#include "Sync/Lock.h"

#include "_Version.h"
#include "Dispatch/Connector.h"
#include "Dispatch/RefVariant.h"
#include "Dispatch/Object.h"
#include "Dispatch/Dispatch.h"
#include "Dispatch/Proxy.h"

#include "Dispatch/Ex/Ex.h"
#include "Dispatch/Ex/Properties.h"
#include "Dispatch/Ex/Proxy.h"

#include "GlobalInterface/GlobalInterface.h"
