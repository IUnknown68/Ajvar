/**************************************************************************//**
 * @file
 * @brief     Declaration of MSIEWindows
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/

#pragma once

namespace Ajvar {

namespace Browser {

#define GET_HWND_PROP(name) \
  __declspec(property(get = get_##name)) HWND name;

//=============================================================================
/// @class   Ajvar::Browser::MSIEWindows
/// @brief   Get the HWNDs of internet explorer related windows.

class MSIEWindows
{
public:
  static HWND GetShellBrowserWindow(ATL::CComQIPtr<IServiceProvider> aServiceProvider)
  {
    if (!aServiceProvider) {
      return nullptr;
    }
    ATL::CComPtr<IOleWindow> window;
    if (FAILED(aServiceProvider->QueryService(SID_SShellBrowser, IID_IOleWindow,(void**)&window))) {
      return nullptr;
    }
    HWND hwndBrowser = nullptr;
    if (FAILED(window->GetWindow(&hwndBrowser))) {
      return nullptr;
    }
    return hwndBrowser;
  }

public:
  MSIEWindows(IUnknown * pUnk)
  {
    SetBrowser(pUnk);
  }

  HRESULT SetBrowser(IUnknown * pUnk)
  {
    mhwndFrame =
    mhwndTabWindow =
    mhwndTab =
    mhwndToolbar =
        nullptr;

    mhwndTabWindow = GetShellBrowserWindow(pUnk);
    if (nullptr == mhwndTabWindow) {
      return E_FAIL;
    }

    mhwndTab = ::GetParent(mhwndTabWindow);
    mhwndFrame = ::GetParent(mhwndTab);
    mhwndToolbar = ::FindWindowEx(mhwndTab, nullptr, nullptr, nullptr);

    return S_OK;
  }

  GET_HWND_PROP(Frame);
  GET_HWND_PROP(TabWindow);
  GET_HWND_PROP(Tab);
  GET_HWND_PROP(Toolbar);

protected:
  HWND mhwndFrame;
  HWND mhwndTabWindow;
  HWND mhwndTab;
  HWND mhwndToolbar;
};

#undef GET_HWND_PROP

} // namespace Browser
} // namespace Ajvar

