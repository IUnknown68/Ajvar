/**************************************************************************//**
@file
@brief     Implementation of `Ajvar::Dispatch::Ex::Properties`.
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#include "stdafx.h"
#include <Dispatch/Ex/Properties.h>

namespace Ajvar {
namespace Dispatch {
namespace Ex {

//----------------------------------------------------------------------------
// CTOR
Properties::Properties(void) :
  mLastDISPID(0)
{
}

//----------------------------------------------------------------------------
// DTOR
Properties::~Properties(void)
{
}

//----------------------------------------------------------------------------
// reset
void Properties::reset()
{
  mNames.clear();
  mValues.clear();
  mNameIDs.clear();
  mLastDISPID = 0;
}

//----------------------------------------------------------------------------
// putValue
HRESULT Properties::putValue(
  DISPID aId,
  const VARIANT & aProperty)
{
  // check if we aId refers to an existing value
  if (mNames.end() == mNames.find(aId)) {
    return DISP_E_MEMBERNOTFOUND;
  }
  return mValues[aId].Copy(&aProperty);
}

//----------------------------------------------------------------------------
// getValue
HRESULT Properties::getValue(
  DISPID aId,
  VARIANT * aRetVal)
{
  if (!aRetVal) {
    return E_POINTER;
  }
  MapDispId2Variant::iterator it = mValues.find(aId);
  if (it == mValues.end()) {
    return DISP_E_MEMBERNOTFOUND;
  }
  return ::VariantCopy(aRetVal, &it->second);
}

//----------------------------------------------------------------------------
// getName
HRESULT Properties::getName(
  DISPID aId,
  BSTR * aRetVal)
{
  if (!aRetVal) {
    return E_POINTER;
  }
  MapDispId2StringOrdered::iterator it = mNames.find(aId);
  if (it == mNames.end()) {
    return DISP_E_MEMBERNOTFOUND;
  }
  return it->second.CopyTo(aRetVal);
}

//----------------------------------------------------------------------------
// getDispID
HRESULT Properties::getDispID(
  LPCOLESTR aName,
  DISPID * aRetVal,
  bool aCreate)
{
  if (!aRetVal) {
    return E_POINTER;
  }
  (*aRetVal) = DISPID_UNKNOWN;

  MapName2DispId::iterator it = mNameIDs.find(aName);
  if (it != mNameIDs.end()) {
    (*aRetVal) = it->second;
    return S_OK;
  }
  if (!aCreate) {
    return DISP_E_UNKNOWNNAME;
  }

  DISPID did = getNextDISPID();
  if (DISPID_UNKNOWN == did) {
    return E_OUTOFMEMORY; // no more IDs available
  }
  mNames[did] = aName;
  (*aRetVal) = mNameIDs[aName] = mLastDISPID = did;

  return S_OK;
}

//----------------------------------------------------------------------------
// enumNextDispID
HRESULT Properties::enumNextDispID(
  DWORD grfdex, /* ignored */
  DISPID did,
  DISPID * aRetVal)
{
  (void)grfdex;
  if (!aRetVal) {
    return E_POINTER;
  }

  MapDispId2StringOrdered::iterator it;
  if (did == DISPID_STARTENUM) {
    // get first entry
    it = mNames.begin();
  }
  else {
    // get current entry
    it = mNames.find(did);
    if (it != mNames.end()) {
      // and progress to the next one
      ++it;
    }
  }

  // As long as we have a valid position...
  while(it != mNames.end()) {
    // ...get the value for this position...
    MapDispId2Variant::iterator itName = mValues.find(it->first);
    if (itName != mValues.end()) {
      // ...and, since we have one, return it.
      (*aRetVal) = itName->first;
      return S_OK;
    }
    // There is no value stored for this id, so progress to next one.
    ++it;
  }
  // sorry, no more data
  return S_FALSE;

}

//----------------------------------------------------------------------------
// remove
HRESULT Properties::remove(DISPID aId)
{
  MapDispId2Variant::iterator it = mValues.find(aId);
  if (it == mValues.end()) {
    return DISP_E_UNKNOWNNAME;
  }
  // The actual value will get erased, name / id will be kept.
  mValues.erase(it);
  return S_OK;
}

//----------------------------------------------------------------------------
// remove
HRESULT Properties::remove(LPCOLESTR aName)
{
  MapName2DispId::iterator it = mNameIDs.find(aName);
  if (it == mNameIDs.end()) {
    return DISP_E_UNKNOWNNAME;
  }
  return remove(it->second);
}

//----------------------------------------------------------------------------
// getNextDISPID
DISPID Properties::getNextDISPID()
{
  DISPID roundTrip = mLastDISPID;
  DISPID did = mLastDISPID + 1;
  // NOTE: We want to go round the full range of DISPID (which is a signed int),
  // so we use "!=" here instead of "<"
  for (did; did != roundTrip; did++) {
    if (DISPID_UNKNOWN == did) {
      continue; // skip DISPID_UNKNOWN
    }
    if (mNames.find(did) == mNames.end()) {
      return did; // does not exist already
    }
  }
  return DISPID_UNKNOWN;  // error: no ID available
}

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar
