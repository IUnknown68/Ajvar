/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch::Ex`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "Connector.h"
#include <functional>

namespace Ajvar {
namespace Dispatch {
namespace Ex {

/// @brief Type for a function passed to `forEach`
/// @details The callback takes the following arguments:
/// - HRESULT aResult : Result from getting the property. Can be an error.
/// - BSTR aName      : Name of the property. May be invalid when `aResult` is an error.
/// - DISPID aDispId  : DISPID of the current value. May be invalid.
/// - VARIANT aValue  : The actual value. May be invalid.
typedef
  std::function<HRESULT(HRESULT aResult, const BSTR & aName, const DISPID aDispId, const VARIANT  & aValue)>
    EachDispEx;

/// @brief Iterate over all properties of `aDispEx` and calls `aEach()` for each found property.
HRESULT forEach(IDispatchEx * aDispEx, EachDispEx aEach, DWORD aEnumFlags = fdexEnumAll|fdexEnumDefault);

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar

