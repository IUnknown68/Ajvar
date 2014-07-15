/**************************************************************************//**
 * @file
 * @brief     Declaration of `ALTX::Dispatch::Ex` Connector class.
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include <DispEx.h>

namespace Ajvar {
namespace Dispatch {
namespace Ex {

/// @brief Our interface type: `IDispatchEx`
typedef IDispatchEx _Interface;

//============================================================================
/// @class  _Connector.
/// @brief  Connects a `IDispatchEx` property to a VARIANT via `Set()` / `Get()`
///         methods.
/// @tparam  Flags Flags for `Set()` or 'Get()', e.g. `fdexNameEnsure`.
template<DWORD Flags>
  class _Connector
{
public:
  /// @brief Our interface type: `IDispatchEx`
  typedef typename _Interface TIf;

  /// @brief Get a `DISPID` for `aName`.
  /// @param aObject  The object to get the property.
  /// @param aName  Name of the property.
  /// @param[retval] aRetVal  Reference to a `DISPID`.
  static HRESULT GetDISPID(IDispatchEx * aObject, LPCWSTR aName, DISPID & aRetVal)
  {
    if (nullptr == aObject || nullptr == aName) {
      return E_INVALIDARG;
    }
	  return aObject->GetDispID(CComBSTR(aName), Flags, &aRetVal);
  }

  /// @brief Get the `VARIANT` for `aDispId` or an error.
  /// @param aObject  The object to get the property.
  /// @param aDispId  'DISPID' of the property.
  /// @param[retval] aRetVal  Reference to a `VARIANT` receiving the property or an error.
  /// @return `HRESULT`. In case of an error `aRetVal` is set to `VT_ERROR` and contains the same `HRESULT`.
  static HRESULT Get(IDispatchEx * aObject, DISPID aDispId, VARIANT & aRetVal)
  {
    if (nullptr == aObject) {
      return E_INVALIDARG;
    }
    // NOTE: We empty the return value here in any case!
    VariantClear(&aRetVal);
	  DISPPARAMS params = {nullptr, nullptr, 0, 0};
	  HRESULT hr = aObject->InvokeEx(
      aDispId,
      LOCALE_USER_DEFAULT,
      DISPATCH_PROPERTYGET,
      &params,
      &aRetVal, NULL, NULL);
    if (FAILED(hr)) {
      // In case of error store the error in return value
      VariantClear(&aRetVal);
      aRetVal.vt = VT_ERROR;
      aRetVal.scode = hr;
    }
    return hr;
  }

  /// @brief Get the `VARIANT` for `aName` or an error.
  /// @see Get(TInterface * aObject, DISPID aDispId, VARIANT & aRetVal)
  /// @param[out] aDispIdRet  If not null, receiving the `DISPID` for `aName`.
  static HRESULT Get(IDispatchEx * aObject, LPCWSTR aName, VARIANT & aRetVal, DISPID * aDispIdRet = nullptr)
  {
    if (nullptr == aObject || nullptr == aName) {
      return E_INVALIDARG;
    }
    DISPID did = 0;
	  HRESULT hr = GetDISPID(aObject, aName, did);
    if (FAILED(hr)) {
      // In case of error store the error in return value
      VariantClear(&aRetVal);
      aRetVal.vt = VT_ERROR;
      aRetVal.scode = hr;
		  return hr;
    }
    if (aDispIdRet) {
      (*aDispIdRet) = did;
    }
    return Get(aObject, did, aRetVal);
  }

  /// @brief Set property `aDispId` to `aValue`.
  /// @param aObject  The object to get the property.
  /// @param aDispId  'DISPID' of the property.
  /// @param aValue The new value.
  static HRESULT Set(IDispatchEx * aObject, DISPID aDispId, VARIANT & aValue)
  {
    if (nullptr == aObject) {
      return E_INVALIDARG;
    }
	  DISPID namedArgs[] = {DISPID_PROPERTYPUT};
	  DISPPARAMS params = {&aValue, namedArgs, 1, 1};
    HRESULT hr = E_FAIL;
	  if (aValue.vt == VT_UNKNOWN || aValue.vt == VT_DISPATCH || 
		  (aValue.vt & VT_ARRAY) || (aValue.vt & VT_BYREF)) {
        // try `DISPATCH_PROPERTYPUTREF` first...
			  hr = aObject->InvokeEx(
          aDispId,
          LOCALE_USER_DEFAULT,
          DISPATCH_PROPERTYPUTREF,
          &params, NULL, NULL, NULL);
	  }
    // if value is a basic type OR the firt `InvokeEx` failed...
    if (FAILED(hr)) {
      // ... try again with `DISPATCH_PROPERTYPUT`
	    hr = aObject->InvokeEx(
        aDispId,
        LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUT,
        &params, NULL, NULL, NULL);
    }
    return hr;
  }

  /// @brief Set property `aName` to `aValue`.
  /// @see Set(TInterface * aObject, DISPID aDispId, VARIANT & aValue)
  static HRESULT Set(IDispatchEx * aObject, LPCWSTR aName, VARIANT & aValue)
  {
    if (nullptr == aObject || nullptr == aName) {
      return E_INVALIDARG;
    }
    DISPID did = 0;
	  HRESULT hr = GetDISPID(aObject, aName, did);
    if (FAILED(hr)) {
		  return hr;
    }
    return Set(aObject, did, aValue);
  }
};

/// @brief  Typedef for `IDispatchEx` connector **creating** new properties.
typedef _Connector<fdexNameEnsure> ConnectorCreate;
  
/// @brief  Typedef for `IDispatchEx` connector **not** creating new properties.
typedef _Connector<0> ConnectorGet;

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar

