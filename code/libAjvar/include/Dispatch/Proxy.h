/**************************************************************************//**
@file
@brief     
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {
namespace Dispatch {

// used by Dispatch::Proxy
#define FORWARD_CALL(_name, ...) \
    { \
      return (mTarget) ? mTarget->_name(__VA_ARGS__) : E_UNEXPECTED;  \
    }

// -------------------------------------------------------------------------
/// @brief Proxy for any `IDispatch` object.
/// @details Holds a pointer to the target object. By default `Proxy`
/// forwards all calls to `mTarget`. `TBaseClass`-methods can be overridden
/// in a derived class. `TBaseClass` can be either `IDispatch` or a class
/// derived from `IDispatch` (typically `IDispatchEx`).
template<class TBaseClass = IDispatch>
  class ATL_NO_VTABLE Proxy :
    public TBaseClass
{
public:
  // -------------------------------------------------------------------------
  // IDispatch methods

  STDMETHOD(GetTypeInfoCount)(UINT *pctinfo)
        FORWARD_CALL(GetTypeInfoCount, pctinfo)
  STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID aLCID, ITypeInfo **ppTInfo)
        FORWARD_CALL(GetTypeInfo, iTInfo, aLCID, ppTInfo)
  STDMETHOD(GetIDsOfNames)(REFIID aRefIID, LPOLESTR *rgszNames, UINT cNames,
      LCID aLCID, DISPID *rgDispId)
        FORWARD_CALL(GetIDsOfNames, aRefIID, rgszNames, cNames, aLCID, rgDispId)
  STDMETHOD(Invoke)(DISPID aDISPID, REFIID aRefIID, LCID aLCID, WORD aFlags,
      DISPPARAMS *aDispParams, VARIANT *aVariantRet,
      EXCEPINFO *aExcepInfo, UINT *aArgumentErrors)
        FORWARD_CALL(Invoke, aDISPID, aRefIID, aLCID, aFlags, aDispParams,
            aVariantRet, aExcepInfo, aArgumentErrors)

  ATL::CComQIPtr<TBaseClass> mTarget;
};

#undef FORWARD_CALL

} //namespace Dispatch
} //namespace Ajvar
