/**************************************************************************//**
 * @file
 * @brief     Main include file for using Ajvar
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

//============================================================================
/// @brief Main namespace: ATL Extensions
/// @details  ATL Extensions inlcude a bunch of helpful stuff for dealing
/// with `IDispatch` and `IDispatchEx`. It uses ATL classes as baseclasses.
namespace Ajvar {

/// @brief Internally used smartVariant
typedef ATL::CComVariant ComVariant;
/// @brief Internally used smartBSTR
typedef ATL::CComBSTR ComBSTR;

} // namespace Ajvar

#include "Dispatch/Dispatch.h"
#include "Dispatch/Ex/Ex.h"

namespace Ajvar {

//============================================================================
/// @brief Namespace for `IDispatch` retlated stuff in ATL Extensions.
namespace Dispatch {

/// @brief `IDispatch` based object.
/// Compatible with CComQIPtr<IDispatch>.
typedef _Object<_Interface, Connector>
  Object;

//============================================================================
/// @brief Namespace for `IDispatchEx` retlated stuff in ATL Extensions.
namespace Ex {

/// @brief `IDispatchEx` based object with autocreating properties.
/// Compatible with CComQIPtr<IDispatchEx>.
typedef _Object<_Interface, ConnectorCreate>
  ObjectCreate;

/// @brief `IDispatchEx` based object **not** autocreating properties.
/// Compatible with CComQIPtr<IDispatchEx>.
typedef _Object<_Interface, ConnectorGet>
  ObjectGet;

/// @brief Default: Create properties.
typedef ObjectCreate
  Object;

} // namespace Ex
} // namespace Dispatch

namespace Time {
/**
 * Ajvar::Time
 * Contains timestamphandling and conversion between
 * Win32 FILETIME based values, unix timestamp and
 * JScript timestamp.
 *
 * NOTE: It is illegal to cast a FILETIME to a time_64, this might produce
 * alignment errors on Win64. That's why the original method was adapted to
 * accept directly a FILETIME structure rather than some int64.
 * See http://msdn.microsoft.com/en-us/library/ms724284%28v=vs.85%29.aspx
 *
 * Ajvar::Time contains the following namespace / type relations:
 *  Win32     : A 64bit uinteger based on the Win32-epoch in 100-nanoseconds
 *  Unix      : A 32bit integer, unix time stamp
 *  JScript   : A double value, JavaScript, as returned from Date.getTime()
 *
 * Each namespace includes:
 * - typedef Time   : See above
 * - Time now()     : Returns now
 * - Time from(...) : methods for conversions
 *
 * Usage:
 *   Win32::Time t = Win32::now();
 *   JScript::Time tJs = JScript::from(t);
 **/

// generic types
// 64bit timestamp
typedef ULONGLONG time_64;

// 32bit timestamp
typedef time_t time_32;

// double timestamp
typedef double time_d;

// constants
// Difference between Win32-epoch and unix-epoch in 100-nanoseconds
static const time_64 EPOCH_DIFF = 116444736000000000LL;

// Factor seconds to milliseconds
static const ULONG RATE_DIFF_S_JS = 1000;

// Factor seconds to 100 nsecs
static const ULONG RATE_DIFF_UNIX = 10000000;

// Factor milliseconds to 100 nsecs
static const ULONG RATE_DIFF_JS = 10000;

// FILETIME2time_64
// Converts a FILETIME to a time_64
inline time_64 FILETIME2time_64(const FILETIME & aFileTime) {
  return (((time_64)aFileTime.dwLowDateTime) | (((time_64)aFileTime.dwHighDateTime) << 32));
}

// Specialized types
namespace Win32 {
  typedef time_64 Time;
}

namespace Unix {
  typedef time_32 Time;
}

namespace JScript {
  typedef time_d Time;
}

// Types: Implementations

// Win32 timestamps
namespace Win32 {
  inline Time from(const Unix::Time & aOther)
  {
    return (Time)aOther * RATE_DIFF_UNIX + EPOCH_DIFF;
  }

  inline Time from(const JScript::Time & aOther)
  {
    return (Time)aOther * RATE_DIFF_JS + EPOCH_DIFF;
  }

  inline Time now()
  {
    FILETIME ft = {0};
    GetSystemTimeAsFileTime(&ft);
    return FILETIME2time_64(ft);
  }
}

// Unix timestamps
namespace Unix {
  inline Time from(const Win32::Time & aOther)
  {
    return (Time)(aOther - EPOCH_DIFF) / RATE_DIFF_UNIX;
  }

  inline Time from(const JScript::Time & aOther)
  {
    return (Time)(aOther / RATE_DIFF_S_JS);
  }

  inline Time now()
  {
    return from(Win32::now());
  }
}

// JScript timestamps
namespace JScript {
  // note that we are flooring incoming integer values.
  inline Time from(const Win32::Time & aOther)
  {
    return (Time)(((Time)aOther - EPOCH_DIFF) / RATE_DIFF_JS);
  }

  inline Time from(const Unix::Time & aOther)
  {
    return (Time)((Time)aOther * RATE_DIFF_S_JS);
  }

  inline Time now()
  {
    return from(Win32::now());
  }
}

} //namespace Time
} // namespace Ajvar
