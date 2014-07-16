/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch::Ex`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "Connector.h"

namespace Ajvar {
namespace Dispatch {

//============================================================================
/// @brief Namespace for `IDispatchEx` retlated stuff in ATL Extensions.
/// @details
/// Classes: `ObjectCreate`, `ObjectGet`, `Object` (=`ObjectGet`)
namespace Ex {

/// @brief Type for a function passed to `forEach`
/// @details The callback takes the following arguments:
/// - HRESULT aResult : Result from getting the property. Can be an error.
/// - BSTR aName      : Name of the property. May be invalid when `aResult` is an error.
/// - DISPID aDispId  : DISPID of the current value. May be invalid when `aResult` is an error.
/// - VARIANT aValue  : The actual value. May be invalid when `aResult` is an error.
typedef
  std::function<HRESULT(HRESULT aResult, const BSTR & aName, const DISPID aDispId, const VARIANT  & aValue)>
    EachDispEx;

/// @brief Iterate over all properties of `aDispEx` and calls `aEach()` for each found property.
HRESULT forEach(IDispatchEx * aDispEx, EachDispEx aEach, DWORD aEnumFlags = fdexEnumAll|fdexEnumDefault);

/// @brief `IDispatchEx` based object with autocreating properties.
/// Compatible with CComQIPtr<IDispatchEx>.
typedef _Object<_Interface, ConnectorCreate>
  ObjectCreate;

/// @brief `IDispatchEx` based object **not** autocreating properties.
/// Compatible with CComQIPtr<IDispatchEx>.
typedef _Object<_Interface, ConnectorGet>
  ObjectGet;

/// @brief Default: Create properties.
typedef ObjectCreate
  Object;


} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar

