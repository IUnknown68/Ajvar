/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch::Ex`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "stdafx.h"
#include <Ajvar.h>

namespace Ajvar {
namespace Dispatch {
namespace Ex {

HRESULT forEach(IDispatchEx * aDispEx, EachDispEx aEach, DWORD aEnumFlags)
{
  DISPID did = DISPID_STARTENUM;
  HRESULT hr = aDispEx->GetNextDispID(aEnumFlags, did, &did);
  DISPPARAMS params = {0};
  while(S_OK == hr) {
    ComBSTR name;
    ComVariant value;

    hr = aDispEx->GetMemberName(did, &name);
    if (SUCCEEDED(hr)) {
      hr = aDispEx->InvokeEx(did, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &value, NULL, NULL);
    }

    hr = aEach(hr, name, did, value);
    if (S_OK != hr) {
      // can also be S_FALSE, which means stop enumerating.
      // Caller will still receive success code.
      return hr;
    }

    hr = aDispEx->GetNextDispID(aEnumFlags, did, &did);
  }

  return S_OK;
}

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar

