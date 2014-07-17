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
- `Ajvar::Time::Timespan`
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

@section Classes Classes
- `Ajvar::Dispatch::Object`
- `Ajvar::Dispatch::RefVariant`
- `Ajvar::Dispatch::Ex::ObjectGet`
- `Ajvar::Dispatch::Ex::ObjectCreate`
- `Ajvar::Dispatch::Ex::Object`
- `Ajvar::Dispatch::Ex::Properties`

@section Templates Templates
- `Ajvar::Dispatch::_Connector`
***************************************************************************/
#pragma once

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

#include "Dispatch/Connector.h"
#include "Dispatch/RefVariant.h"
#include "Dispatch/Object.h"
#include "Dispatch/Dispatch.h"
#include "Dispatch/Ex/Ex.h"
#include "Dispatch/Ex/Properties.h"
#include "Ajvartime.h"

