/**************************************************************************//**
 * @file
 * @brief     Declaration of `ALTX::Dispatch::RefVariant`.
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

namespace Ajvar {
namespace Dispatch {

//============================================================================
/// @class  Ajvar::Dispatch::RefVariant
/// @brief  For using a `CComVariant`-based property as an lvalue.
/// @see    _Object
/// @details A `RefVariant` is returned from `_Object` when
/// querying a property. It can be used as lvalue in an assignment. It keeps a reference
/// to the `_Object` (usually `IDispatch` or `IDispatchEx`) it belongs to and sets
/// the value back on the interface when the assignment operator is called.
/// Additionally it is a callable, in case the property is a function.
/// It uses a `TConnector` to get / set the property.
/// @tparam  TConnector The connector class to use. @See Ajvar::Dispatch::_Connector.
template<class TConnector, class TBase = ComVariant>
  class RefVariant :
    public TBase
{
public:
  /// @brief Our interface (usually `IDispatch` or `IDispatchEx`)
  typedef typename TConnector::TIf TIf;

  /// @brief Contructor taking a reference to the the object this property
  /// belongs to.
  /// @param[in]  ref Reference to the object this property originates from.
  RefVariant(TIf * aObject, LPCWSTR aName) :
    mObject(aObject), mDispId(0)
  {
    TConnector::Get(mObject, aName, *this, &mDispId);
  }

  /// @{
  /// @brief Assignment operators.
  /// @details  Does assignment via `TBase`, then calls
  /// `Set()` to set this property on the underlying object.
  /// We tried using a template function, here, but it does not work
  /// reliably because `CComVariant` (the default baseclass) already
  /// declares these operators.
#define ASSIGNMENT(_type) \
  RefVariant & \
    operator = (_type aValue) ATLVARIANT_THROW() \
  { \
    TBase::operator =(aValue); \
    TConnector::Set(mObject, mDispId, *this); \
    return *this; \
  }

  ASSIGNMENT(const RefVariant&)
  ASSIGNMENT(const ComVariant&)
  ASSIGNMENT(const VARIANT&)
  ASSIGNMENT(const ATL::CComBSTR&)
  ASSIGNMENT(LPCOLESTR)
  ASSIGNMENT(LPCSTR)
  ASSIGNMENT(bool)
  ASSIGNMENT(int)
  ASSIGNMENT(BYTE)
  ASSIGNMENT(short)
  ASSIGNMENT(long)
  ASSIGNMENT(float)
  ASSIGNMENT(double)
  ASSIGNMENT(CY)
  ASSIGNMENT(IDispatch*)
  ASSIGNMENT(IUnknown*)
  ASSIGNMENT(char)
  ASSIGNMENT(unsigned short)
  ASSIGNMENT(unsigned long)
  ASSIGNMENT(unsigned int)
  ASSIGNMENT(BYTE*)
  ASSIGNMENT(short*)
#ifdef _NATIVE_WCHAR_T_DEFINED
  ASSIGNMENT(USHORT*)
#endif
  ASSIGNMENT(int*)
  ASSIGNMENT(UINT*)
  ASSIGNMENT(long*)
  ASSIGNMENT(ULONG*)
#if (_WIN32_WINNT >= 0x0501) || defined(_ATL_SUPPORT_VT_I8)
  ASSIGNMENT(LONGLONG)
  ASSIGNMENT(LONGLONG*)
  ASSIGNMENT(ULONGLONG)
  ASSIGNMENT(ULONGLONG*)
#endif
  ASSIGNMENT(float*)
  ASSIGNMENT(double*)
  ASSIGNMENT(const SAFEARRAY *)

#undef ASSIGNMENT
  /// @}


  /// @brief Call operator.
  /// @details  Takes up to 4 optional arguments for the call.
  /// The arguments here are in the order as they appear in the callee.
  /// @return   The return value of the call, or an error.
#define CALL_ARGUMENT_COUNT 4
  ComVariant operator () (
    ComVariant aArg0 = ComVariant(),
    ComVariant aArg1 = ComVariant(),
    ComVariant aArg2 = ComVariant(),
    ComVariant aArg3 = ComVariant())
  {
    ComVariant retval;
    if (!mObject) {
      retval.vt = VT_ERROR;
      retval.scode = E_UNEXPECTED;
      return retval;
    }
    // add arguments in reversed order
    VARIANT args[CALL_ARGUMENT_COUNT] = {aArg3, aArg2, aArg1, aArg0};
    UINT firstArg = 0;  // all arguments are valid
    for (firstArg; (firstArg < CALL_ARGUMENT_COUNT) && (VT_EMPTY == args[firstArg].vt); firstArg++);

    HRESULT hr = call(retval, args + firstArg, CALL_ARGUMENT_COUNT - firstArg);
    if (FAILED(hr)) {
      retval.vt = VT_ERROR;
      retval.scode = hr;
    }
    return retval;
  }
#undef CALL_ARGUMENT_COUNT

  /// @brief Lowlevel call method.
  /// @param aRetVal Return value of the call, or an error.
  /// @details  In case this instance refers to a function, call the function.
  /// This is the lowlevel method, taking a raw `VARIANT` pointer to the arguments,
  /// and the number of arguments in this array.
  ///
  /// As usual the values are in reversed order.
  HRESULT call(
    ComVariant & aRetVal,
    VARIANT * aArgs,
    UINT nArgCount)
  {
    if (!mObject) {
      return E_UNEXPECTED;
    }
    DISPPARAMS params = {aArgs, nullptr, nArgCount, 0};
    return mObject->InvokeEx(
      mDispId,
      LOCALE_USER_DEFAULT,
      DISPATCH_METHOD,
      &params,
      &aRetVal, NULL, NULL);
  }

protected:
  /// @brief Reference to the connected object.
  ATL::CComPtr<TIf> mObject;

  /// @brief `DISPID` for the property this instance represents.
  DISPID mDispId;
};

} // namespace Dispatch
} // namespace Ajvar
