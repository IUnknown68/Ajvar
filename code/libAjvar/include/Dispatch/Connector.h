/**************************************************************************//**
 * @file
 * @brief     Declaration of `ALTX::Dispatch` Connector class.
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

namespace Ajvar {
namespace Dispatch {

/// @brief Our interface type: `IDispatch`
typedef IDispatch _Interface;

//============================================================================
/// @class  _Connector.
/// @brief  Connects a `IDispatch` property to a VARIANT via `Set()` / `Get()`
///         methods.
/// @tparam  Flags Flags for `Set()` or 'Get()'. Not used in `IDispatch`.
template<DWORD Flags>
  class _Connector
{
public:
  /// @brief Our interface type: `IDispatch`
  typedef typename _Interface TIf;

  /// @brief Get a `DISPID` for `aName`.
  /// @param aObject  The object to get the property.
  /// @param aName  Name of the property.
  /// @param[retval] aRetVal  Reference to a `DISPID`.
  static HRESULT GetDISPID(IDispatch * aObject, LPCWSTR aName, DISPID & aRetVal)
  {
    if (nullptr == aObject || nullptr == aName) {
      return E_INVALIDARG;
    }
    return aObject->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&aName), 1, LOCALE_USER_DEFAULT, &aRetVal);
  }

  /// @brief Get the `VARIANT` for `aDispId` or an error.
  /// @param aObject  The object to get the property.
  /// @param aDispId  'DISPID' of the property.
  /// @param[retval] aRetVal  Reference to a `VARIANT` receiving the property or an error.
  /// @return `HRESULT`. In case of an error `aRetVal` is set to `VT_ERROR` and contains the same `HRESULT`.
  static HRESULT Get(IDispatch * aObject, DISPID aDispId, VARIANT & aRetVal)
  {
    // NOTE: We empty the return value here in any case!
    VariantClear(&aRetVal);
    return ATL::CComPtr<IDispatch>::GetProperty(aObject, aDispId, &aRetVal);
  }

  /// @brief Get the `VARIANT` for `aName` or an error.
  /// @see Get(TInterface * aObject, DISPID aDispId, VARIANT & aRetVal)
  /// @param[out] aDispIdRet  If not null, receiving the `DISPID` for `aName`.
  static HRESULT Get(IDispatch * aObject, LPCWSTR aName, VARIANT & aRetVal, DISPID * aDispIdRet = nullptr)
  {
    if (nullptr == aObject || nullptr == aName) {
      return E_INVALIDARG;
    }
    DISPID did = 0;
	  HRESULT hr = GetDISPID(aObject, aName, did);
    if (FAILED(hr)) {
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
  static HRESULT Set(IDispatch * aObject, DISPID aDispId, VARIANT & aValue)
  {
    return ATL::CComPtr<IDispatch>::PutProperty(aObject, aDispId, &aValue);
  }

  /// @brief Set property `aName` to `aValue`.
  /// @see Set(TInterface * aObject, DISPID aDispId, VARIANT & aValue)
  static HRESULT Set(IDispatch * aObject, LPCWSTR aName, VARIANT & aValue)
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

/// @brief  Connector for `IDispatch`.
typedef _Connector<0> Connector;

} // namespace Dispatch
} // namespace Ajvar

