// Logger.h : Declaration of the Logger

#pragma once
#include <map>

#include "Ajvar_i.h"

namespace Ajvar {
namespace Dispatch {
namespace Ex {

class ATL_NO_VTABLE ObjectImpl :
  public Properties,
  public Ajvar::Debug::AjvarComObjectRootEx<ObjectImpl, ATL::CComSingleThreadModel>,
  public IJSObject,
	public IDispatchEx
{
public:
  typedef ATL::CComObject<ObjectImpl>  _ComObject;

  //--------------------------------------------------------------------------
  // static creator
  static ATL::CComPtr<IJSObject> createInstance(HRESULT & hr);

  ObjectImpl();

  DECLARE_NO_REGISTRY()
	DECLARE_PROTECT_FINAL_CONSTRUCT()
  BEGIN_COM_MAP(ObjectImpl)
	  COM_INTERFACE_ENTRY(IJSObject)
	  COM_INTERFACE_ENTRY(IDispatch)
	  COM_INTERFACE_ENTRY(IDispatchEx)
  END_COM_MAP()

	HRESULT FinalConstruct();

  void FinalRelease();

public:
  // -------------------------------------------------------------------------
  // IDispatch methods
  STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);

  STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);

  STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames,
                                         LCID lcid, DISPID *rgDispId);

  STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
                                  DISPPARAMS *pDispParams, VARIANT *pVarResult,
                                  EXCEPINFO *pExcepInfo, UINT *puArgErr);


  // -------------------------------------------------------------------------
  // IDispatchEx methods
  STDMETHOD(GetDispID)(BSTR bstrName, DWORD grfdex, DISPID *pid);

  STDMETHOD(InvokeEx)(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp,
                      VARIANT *pvarRes, EXCEPINFO *pei,
                      IServiceProvider *pspCaller);

  STDMETHOD(DeleteMemberByName)(BSTR bstrName,DWORD grfdex);

  STDMETHOD(DeleteMemberByDispID)(DISPID id);

  STDMETHOD(GetMemberProperties)(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);

  STDMETHOD(GetMemberName)(DISPID id, BSTR *pbstrName);

  STDMETHOD(GetNextDispID)(DWORD grfdex, DISPID id, DISPID *pid);

  STDMETHOD(GetNameSpaceParent)(IUnknown **ppunk);

private:
  HRESULT dispatchMethod(DISPID dispIdMember, LCID lcid,
                    DISPPARAMS *pDispParams, VARIANT *pVarResult,
                    EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller);

  HRESULT dispatchPropertyGet(WORD wFlags, DISPID dispIdMember, REFIID riid,
                    LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                    EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller);

  HRESULT dispatchPropertyPut(WORD wFlags, DISPID dispIdMember, REFIID riid,
                    LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                    EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller);

  HRESULT dispatchConstruct(DISPID dispIdMember, REFIID riid,
                    LCID lcid, DISPPARAMS *pDispParams, VARIANT *pVarResult,
                    EXCEPINFO *pExcepInfo, IServiceProvider *pspCaller);

};

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar

