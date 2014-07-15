/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "Connector.h"
#include "RefVariant.h"
#include "Object.h"

namespace Ajvar {
namespace Dispatch {


/// @brief Function to get the scripting interface from an `IWebBrowser2`.
/// @details  This gets the current document from an `IWebBrowser2`, from
/// the document the `parentWindow` (if any) and does `QueryInterface()` to get
/// the `TInterface`.
/// @param[in]  aBrowser The webbrowser control.
/// @retval aRetVal The resulting `TInterface`.
template<class TInterface> HRESULT GetScriptDispatch(::IWebBrowser2 * aBrowser, TInterface ** aRetVal)
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

} // namespace Dispatch
} // namespace Ajvar

