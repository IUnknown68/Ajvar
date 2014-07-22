/**************************************************************************//**
@file
@brief     
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {
namespace Dispatch {
namespace Ex {

// used by DispatchEx::Proxy
#define FORWARD_CALL(_name, ...) \
    { \
      return (mTarget) ? mTarget->_name(__VA_ARGS__) : E_UNEXPECTED;  \
    }

// -------------------------------------------------------------------------
/// @brief Proxy for any `IDispatchEx` object.
/// @see `Ajvar::Dispatch::Proxy`
class ATL_NO_VTABLE Proxy :
  public Ajvar::Dispatch::Proxy<IDispatchEx>
{
public:
  // -------------------------------------------------------------------------
  // IDispatchEx methods
  STDMETHOD(GetDispID)(BSTR bstrName, DWORD grfdex, DISPID *pid)
        FORWARD_CALL(GetDispID, bstrName, grfdex, pid)
  STDMETHOD(InvokeEx)(DISPID aDISPID, LCID aLCID, WORD aFlags, DISPPARAMS *aDispParams,
      VARIANT *aVariantRet, EXCEPINFO *aExcepInfo,
      IServiceProvider *aServiceProvider)
        FORWARD_CALL(InvokeEx, aDISPID, aLCID, aFlags, aDispParams, aVariantRet,
           aExcepInfo, aServiceProvider)
  STDMETHOD(DeleteMemberByName)(BSTR bstrName,DWORD grfdex)
        FORWARD_CALL(DeleteMemberByName, bstrName, grfdex)
  STDMETHOD(DeleteMemberByDispID)(DISPID aDISPID)
        FORWARD_CALL(DeleteMemberByDispID, aDISPID)
  STDMETHOD(GetMemberProperties)(DISPID aDISPID, DWORD grfdexFetch, DWORD *pgrfdex)
        FORWARD_CALL(GetMemberProperties, aDISPID, grfdexFetch, pgrfdex)
  STDMETHOD(GetMemberName)(DISPID aDISPID, BSTR *pbstrName)
        FORWARD_CALL(GetMemberName, aDISPID, pbstrName)
  STDMETHOD(GetNextDispID)(DWORD grfdex, DISPID aDISPID, DISPID *pid)
        FORWARD_CALL(GetNextDispID, grfdex, aDISPID, pid)
  STDMETHOD(GetNameSpaceParent)(IUnknown **ppunk)
        FORWARD_CALL(GetNameSpaceParent, ppunk)
};
#undef FORWARD_CALL

} //namespace Ex
} //namespace Dispatch
} //namespace Ajvar
