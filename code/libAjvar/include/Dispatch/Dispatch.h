/**************************************************************************//**
@file
@brief     Main include file for `Ajvar::Dispatch`
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

#include "_Object.h"

namespace Ajvar {

//============================================================================
/// @brief Namespace for `IDispatch` related stuff in ATL Extensions.
namespace Dispatch {

/// @brief Function to get the scripting interface from an `IWebBrowser2`.
/// @details Gets the current `IHTMLDocument2` from an `IWebBrowser2` via
/// the `Document` propert, then gets the `parentWindow` from it, and finally
/// does a `QueryInterface()` to get the `TInterface`.
/// @param[in]  aBrowser The webbrowser control.
/// @retval aRetVal The resulting `TInterface`.
template<class TInterface>
  HRESULT GetScriptDispatch(::IWebBrowser2 * aBrowser, TInterface ** aRetVal)
{
  if (!aBrowser) {
    return E_INVALIDARG;
  }
  if (!aRetVal) {
    return E_POINTER;
  }
  (*aRetVal) = nullptr;

  CComPtr<IDispatch> documentDispatch;
  HRESULT hr = aBrowser->get_Document(&documentDispatch);
  if (FAILED(hr)) {
    return hr;
  }

  CComQIPtr<IHTMLDocument2> htmlDocument(documentDispatch);
  if (nullptr == htmlDocument) {
    return E_NOINTERFACE;
  }

  CComPtr<IHTMLWindow2> htmlWindow;
  hr = htmlDocument->get_parentWindow(&htmlWindow);
  if (FAILED(hr)) {
    return hr;
  }

  return htmlWindow->QueryInterface(_uuidof(TInterface), (void **)aRetVal);
}

/// @brief `IDispatch` based object.
/// @details This is the class you will actually use.
/// Compatible with `CComQIPtr<IDispatch>`.
typedef _Object<_Interface, Connector>
  Object;


} // namespace Dispatch
} // namespace Ajvar

