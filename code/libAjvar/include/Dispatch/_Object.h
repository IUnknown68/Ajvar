/**************************************************************************//**
@file
@brief     Declaration of `Ajvar::Dispatch::Represents`.
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

#include "Connector.h"
#include "RefVariant.h"

namespace Ajvar {
namespace Dispatch {

//============================================================================
/// @class  _Object
/// @brief  Represents an object with a `IDispatch` or `IDispatchEx` interface.
/// @details This class represents an object, means, a COM object implementing
/// `IDispatch` or `IDispatchEx`. Usually you would use it with script objects.
///
/// To do so it:
/// - Declares an array access operator `[]` for easy access of the object's
/// properties.
/// - Derives from `CComQIPtr<TInterface>` and has additional assignment
/// operators and constructors for getting the script interface
/// from an `IWebBrowser2` interface.
/// - has a `constructNew` method to create a new `Object`, `Array` or any
/// other object using `IDispatch`'s construct mechanism.
///
/// Sample:
/// @code{.cpp}
/// using namespace Ajvar;
///
/// CComPtr<IWebBrowser2> browser;  // from somewhere
///
/// // get the scripting (htmlwindow) from browser
/// Dispatch::Ex::Object script(browser);
///
/// // get a value
/// auto foo = script[L"foo"];
///
/// // 'foo' is now a Ajvar::Dispatch::RefVariant.
///
/// // Can check now for error and correct type:
/// if (VT_ERROR == foo.vt) ....
/// if (VT_I4 != foo.vt) ....
///
/// // assign a new value to "foo"
/// foo = foo.lVal * 3;
///
/// // in JS 'window.foo' will be changed now.
///
/// // create a new Object
/// auto message = script.constructNew(L"Object");
/// // 'message' is now a Dispatch::Ex::Object
///
/// message[L"text"] = L"Hello World";
/// message[L"document"] = script[L"document"];
///
/// // and set...
/// script[L"message"] = message;
///
/// // ... or call a function
/// script[L"showMessage"]((IDispatchEx*)message, L"Finito");
/// @endcode
template<typename TInterface, class TConnector>
  class _Object :
    public ATL::CComQIPtr<TInterface>
{
public:
  typedef ATL::CComQIPtr<TInterface> _Base;
  typedef RefVariant<TConnector> _LVariant;

public:
  /// @{
  /// @brief Constructors from baseclass.
  _Object() throw() :
    _Base()
  {
  }

  _Object(decltype(__nullptr)) throw()
  {
  }

  _Object(_Inout_opt_ TInterface * lp) throw() :
    _Base(lp)
  {
  }

  _Object(_Inout_ const _Object<TInterface, TConnector>& lp) throw() :
    _Base(lp.p)
  {
  }

  _Object(_Inout_opt_ IUnknown* lp) throw() :
    _Base(lp)
  {
  }
  /// @}

  /// @brief Constructor taking a webbrowser control
  /// @details  Accepts a `IWebBrowser2` interface and gets the script dispatch
  /// from it.
  /// @param[in]  lp Raw `IWebBrowser2` pointer.
  _Object(_Inout_opt_ ::IWebBrowser2 * lp) throw()
  {
    p = nullptr;
    GetScriptDispatch(lp, &p);
  }

  /// @brief Constructor taking a `VARIANT`
  /// @details  If the `VARIANT` is a `VT_DISPATCH` and not null, query for
  /// the `TInterface` interface.
  /// @param[in]  vt `const VARIANT &` to get the `TInterface` from.
  _Object(_Inout_opt_ const VARIANT & vt) throw()
  {
    if ( (VT_DISPATCH == vt.vt)
      && (nullptr != vt.pdispVal)
      && FAILED(vt.pdispVal->QueryInterface(_uuidof(TInterface), (void **)&p))) {
        p = nullptr;
    }
  }

  /// @brief Assignment operator taking a webbrowser control.
  /// @details  Accepts a `IWebBrowser2` interface and gets the script dispatch
  /// from it.
  /// @param[in]  lp Raw `IWebBrowser2` pointer.
  TInterface * operator=(_Inout_opt_ ::IWebBrowser2* lp) throw()
  {
    _Base newScript;
    GetScriptDispatch(lp, &newScript.p);
    return (*this) = newScript.p;
  }

  /// @brief Array access operator taking a property name.
  /// @details  Takes a property name and returns an `_LVariant`
  /// for that property.
  /// @param[in]  aName Name of the property.
  /// @return  `_LVariant` that can be used as lvalue in assingment.
  _LVariant operator [] (LPCWSTR aName)
  {
    return _LVariant(this->p, aName);
  }

  /// @brief Construct a new instance of type `aClass` (usually `Object` or `Array`)
  /// @param[in]  aClass Classname - e.g. "Object" or "Array".
  /// @return  `_Object` holding the new instance.
  _Object<TInterface, TConnector> constructNew(LPCWSTR aClass)
  {
    if ((nullptr == p) || (nullptr == aClass) ) {
      return nullptr;
    }

    // get dispid for classname
    DISPID dispid = 0;
    HRESULT hr = p->GetDispID(CComBSTR(aClass), 0, &dispid);
    if (FAILED(hr)) {
      return nullptr;
    }

    // call constructor
    ATL::CComVariant vt;
    DISPPARAMS params = {nullptr, nullptr, 0, 0};
    hr = p->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_CONSTRUCT,
      &params, &vt, nullptr, nullptr);
    if (FAILED(hr)) {
      return nullptr;
    }
    return vt.pdispVal;
  }
};

} // namespace Dispatch
} // namespace Ajvar
