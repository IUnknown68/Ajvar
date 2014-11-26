/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch::Ex`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "stdafx.h"
#include <Ajvar.h>
#include "Ajvar_i.c"

namespace Ajvar {
namespace Dispatch {
namespace Ex {

//--------------------------------------------------------------------------
// static creator
ATL::CComPtr<IJSObject> ObjectImpl::createInstance(HRESULT & hr)
{
  _ComObject * newInstance = NULL;
  hr = _ComObject::CreateInstance(&newInstance);
  if (SUCCEEDED(hr)) {
    ATL::CComPtr<IJSObject> owner(newInstance);
    return owner;
  }
  return NULL;
}

ObjectImpl::ObjectImpl()
{
}

HRESULT ObjectImpl::FinalConstruct()
{
  return S_OK;
}

void ObjectImpl::FinalRelease()
{
}

// -------------------------------------------------------------------------
// IDispatch methods
STDMETHODIMP ObjectImpl::GetTypeInfoCount(UINT *pctinfo)
{
  IF_NULL_RET(pctinfo);
  (*pctinfo) = 0;
  return S_OK;
}

STDMETHODIMP ObjectImpl::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
  return E_NOTIMPL;
}

STDMETHODIMP ObjectImpl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames,
                                        LCID lcid, DISPID *rgDispId)
{
  return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP ObjectImpl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                DISPPARAMS *pDispParams, VARIANT *pVarResult,
                                EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
  return DISP_E_MEMBERNOTFOUND;
}


// -------------------------------------------------------------------------
// IDispatchEx methods
STDMETHODIMP ObjectImpl::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
  return getDispID(bstrName, pid, (grfdex & fdexNameEnsure) != 0);
}

STDMETHODIMP ObjectImpl::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp,
                    VARIANT *pvarRes, EXCEPINFO *pei,
                    IServiceProvider *pspCaller)
{
  switch(wFlags) {
    case DISPATCH_METHOD:
      return dispatchMethod(id, lcid, pdp, pvarRes, pei, pspCaller);
    case DISPATCH_PROPERTYGET:
      return dispatchPropertyGet(wFlags, id, IID_NULL, lcid, pdp, pvarRes,
        pei, pspCaller);
    case DISPATCH_METHOD|DISPATCH_PROPERTYGET:
      // http://msdn.microsoft.com/en-us/library/windows/desktop/ms221486%28v=vs.85%29.aspx :
      // "Some languages cannot distinguish between retrieving a property and
      // calling a method. In this case, you should set the flags
      // DISPATCH_PROPERTYGET and DISPATCH_METHOD".
      // This means, we have to handle both cases depending on the actual type.
      // If it is callable call it, otherwise return the property
      {
        HRESULT hr = dispatchMethod(id, lcid, pdp, pvarRes, pei, pspCaller);
        if (DISP_E_TYPEMISMATCH == hr) {
          return dispatchPropertyGet(wFlags, id, IID_NULL, lcid, pdp, pvarRes,
            pei, pspCaller);
        }
        return hr;
      }
    case DISPATCH_PROPERTYPUT:
    case DISPATCH_PROPERTYPUTREF:
    case DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF:
      return dispatchPropertyPut(wFlags, id, IID_NULL, lcid, pdp, pvarRes,
        pei, pspCaller);
    case DISPATCH_CONSTRUCT:
      return dispatchConstruct(id, IID_NULL, lcid, pdp, pvarRes, pei, pspCaller);
  }

  ATLTRACE(_T("ObjectImpl::InvokeEx ************* Unknown Flag 0x%08x\n"), wFlags);
  ATLASSERT(0 && "Ooops - did we forget to handle something?");
  return E_INVALIDARG;
}

STDMETHODIMP ObjectImpl::DeleteMemberByName(BSTR bstrName,DWORD grfdex)
{
  return remove(bstrName);
}

STDMETHODIMP ObjectImpl::DeleteMemberByDispID(DISPID id)
{
  return remove(id);
}

STDMETHODIMP ObjectImpl::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
  return E_FAIL;
}

STDMETHODIMP ObjectImpl::GetMemberName(DISPID id, BSTR *pbstrName)
{
  return getName(id, pbstrName);
}

STDMETHODIMP ObjectImpl::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
  return enumNextDispID(grfdex, id, pid);
}

STDMETHODIMP ObjectImpl::GetNameSpaceParent(IUnknown **ppunk)
{
  InternalAddRef();
  return E_FAIL;
}

HRESULT ObjectImpl::dispatchMethod(DISPID dispIdMember, LCID lcid,
                  DISPPARAMS *pDispParams, VARIANT *pVarResult,
                  EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller)
{
  auto it = mValues.find(dispIdMember);
  if (it != mValues.end()) {
    if (it->second.vt != VT_DISPATCH) {
      return DISP_E_TYPEMISMATCH;
    }

    // check if there is already a "this" arg in pDispParams
    bool hasThis = false;
    UINT namedArgCount = pDispParams->cNamedArgs;
    for (UINT n = 0; n < namedArgCount; n++) {
      if (DISPID_THIS == pDispParams->rgdispidNamedArgs[n]) {
        hasThis = true;
        break;
      }
    }

    DISPPARAMS params = {pDispParams->rgvarg, pDispParams->rgdispidNamedArgs, pDispParams->cArgs, pDispParams->cNamedArgs};
    ATL::CComQIPtr<IDispatchEx> dispEx(it->second.pdispVal);

    // if there is no "this", set ourselfs
    if (!hasThis) {
      UINT argCount = pDispParams->cArgs;

      // add "this"
      argCount++;
      std::vector<ATL::CComVariant> args(argCount);
      args[0] = this;
      for (UINT n = 1; n < argCount; n++) {
        args[n] = pDispParams->rgvarg[n-1];
      }

      // add DISPID_THIS
      namedArgCount++;
      std::vector<DISPID> namedArgs(namedArgCount);
      namedArgs[0] = DISPID_THIS;
      for (UINT n = 1; n < namedArgCount; n++) {
        args[n] = pDispParams->rgdispidNamedArgs[n-1];
      }

      DISPPARAMS params = {args.data(), namedArgs.data(), argCount, namedArgCount};

      // Use InvokeEx if the dispatch implements IDispatchEx.
      if (dispEx) {
        return dispEx->InvokeEx(0, lcid, DISPATCH_METHOD,
            &params, pVarResult, pExcepInfo, pspCaller);
      }
      else {
        return it->second.pdispVal->Invoke(0, IID_NULL, lcid, DISPATCH_METHOD,
            &params, pVarResult, pExcepInfo, NULL);
      }
    }

    // Use InvokeEx if the dispatch implements IDispatchEx.
    if (dispEx) {
      return dispEx->InvokeEx(0, lcid, DISPATCH_METHOD,
          pDispParams, pVarResult, pExcepInfo, pspCaller);
    }
    else {
      return it->second.pdispVal->Invoke(0, IID_NULL, lcid, DISPATCH_METHOD,
          pDispParams, pVarResult, pExcepInfo, NULL);
    }
  }

  return DISP_E_MEMBERNOTFOUND;
}

HRESULT ObjectImpl::dispatchPropertyGet(WORD wFlags, DISPID dispIdMember, REFIID riid,
                  LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                  EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller)
{
  return getValue(dispIdMember, pVarResult);
}

HRESULT ObjectImpl::dispatchPropertyPut(WORD wFlags, DISPID dispIdMember, REFIID riid,
                  LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                  EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller)

{
  if (pDispParams->cArgs < 1) {
    return E_INVALIDARG;
  }
  return putValue(dispIdMember, pDispParams->rgvarg[0]);
}

HRESULT ObjectImpl::dispatchConstruct(DISPID dispIdMember, REFIID riid,
                  LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                  EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller)
{
  return E_NOTIMPL;
}



} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar