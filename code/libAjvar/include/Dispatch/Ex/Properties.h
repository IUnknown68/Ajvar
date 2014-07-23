/**************************************************************************//**
@file
@brief     Declaration of `Ajvar::Dispatch::Ex::Properties`.
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {
namespace Dispatch {
namespace Ex {

//============================================================================
/// @class  Properties
/// @brief  Implements expando properties for `IDispatchEx`.
/// @details This class is supposed to be used whenever you need a full-fledged
/// `IDispatchEx` implementation. It does not implement the interface itself,
/// but the logic required to serve the interface.
class Properties
{
public: // types
  /// @brief  Map string => `DISPID`
  typedef std::unordered_map<std::wstring, DISPID>      MapName2DispId;

  /// @brief  Map `DISPID` => `CComVariant`. Contains the actual values.
  typedef std::unordered_map<DISPID, ATL::CComVariant>  MapDispId2Variant;

  /// @brief  Map `DISPID` => `CComBSTR`. Reverse-lookup of names.
  /// Used for enumeration, so it is ordered.
  typedef std::map<DISPID, ATL::CComBSTR>               MapDispId2StringOrdered;

public: // methods

  /// @{
  /// @brief Constructors / destructor.
  Properties(void);
  virtual ~Properties(void);
  /// @}

protected:  // methods
  /// @brief Removes all properties and resets `mLastDISPID`.
  void reset();

  /// @brief Sets a value by id.
  /// @details Sets a value for which a name/id pair has to exist already. To create
  /// such a pair call `getDispID(...)` with `aCreate` set to `true`.
  /// @param[in] aId `DISPID` of the value to set.
  /// @param[in] aProperty `VARIANT *`, the actual value.
  HRESULT putValue(
    DISPID aId,
    const VARIANT & aProperty);

  /// @brief Gets a value by id.
  /// @param[in] aId `DISPID` of the value to get.
  /// @param[out] aRetVal `VARIANT *`, receives the current value.
  HRESULT getValue(
    DISPID aId,
    VARIANT * aRetVal);

  /// @brief Gets the name for a certain id.
  /// @param[in] aId `DISPID` to get the name for.
  /// @param[out] aRetVal `BSTR *`, receives the name.
  HRESULT getName(
    DISPID aId,
    BSTR * aRetVal);

  /// @brief Gets / creates an id for a name.
  /// @details If `aCreate` is true, the property will be created, but no value will be set.
  /// @param[in] aName `LPCOLESTR` to get the `DISPID` for.
  /// @param[out] aRetVal `DISPID`, receives the id.
  /// @param[in] aCreate `bool` - create the property or not.
  HRESULT getDispID(
    LPCOLESTR aName,
    DISPID * aRetVal,
    bool aCreate = false);

  /// @brief Gets the next used (non-empty) `DISPID` for a given id for enumerating
  /// properties.
  /// @details `enumNextDispID()` enumerates the ids of all currently existing
  /// values. This means, deleted or not-yet set values will be skipped, so you
  /// can expect non-continuous ids for subsequent calls.
  /// @code{.cpp}
  /// DISPID did = DISPID_STARTENUM;
  /// for (HRESULT hr = enumNextDispID(0, did, &did);
  ///     S_OK == hr;
  ///     hr = enumNextDispID(0, did, &did)) {
  ///   // do something with did here
  /// }
  /// @endcode
  /// @param[in] grfdex `DWORD` - ignored.
  /// @param[in] did `DISPID`, current `DISPID` to start the enumeration.
  /// @param[out] aRetVal `DISPID` - receives the `DISPID` of the next value.
  /// @return `S_FALSE` if there are no more ids, `S_OK` for success, any other
  /// value on error.
  HRESULT enumNextDispID(
    DWORD grfdex, /* ignored */
    DISPID did,
    DISPID * aRetVal);

  /// @brief Removes a certain property by id.
  /// @details Values will be removed, but the name / id will be kept.
  /// `IDispatchEx` requires that the name/DISPID combination stays valid for the
  /// whole lifetime of the object.
  /// @param[in] aId `DISPID` of the value to remove.
  HRESULT remove(DISPID aId);

  /// @brief Removes a certain property by name.
  /// @see `remove(DISPID aId)`
  /// @param[in] aName `LPCOLESTR`, name of the value to remove.
  HRESULT remove(LPCOLESTR aName);

  /// @brief Gets the next available id.
  DISPID getNextDISPID();

protected:  // attributes
  /// @brief Ordered map containing ids and names.
  MapDispId2StringOrdered mNames; // key: DISPID, value: name

  /// @brief Map containing ids and values.
  MapDispId2Variant mValues;      // key: DISPID, value: property value

  /// @brief Map containing names and ids.
  MapName2DispId mNameIDs;        // key: name, value: DISPID

  /// @brief The last used `DISPID`.
  DISPID mLastDISPID;
};

} // namespace Ex
} // namespace Dispatch
} // namespace Ajvar
