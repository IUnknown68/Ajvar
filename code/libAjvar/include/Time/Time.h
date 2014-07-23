/**************************************************************************//**
@file
@brief     Ajvar::Time namespace
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {

/// @brief Contains conversion between Win32 FILETIME based values, unix
/// timestamp and JScript timestamp.
namespace Time {

/// @brief Compile-time calculated time constants
/// @details
/// `Timespan` contains an enum with `seconds`, `minutes`, `hours`, `days`.
/// So instead of using magic numbers use:
/// @code{.cpp}
/// // setting an interval of one day in seconds:
/// SetInterval(Day::seconds);
///
/// // setting an interval of one day in hours:
/// SetInterval(Day::hours);
///
/// // setting an interval of one year in days:
/// SetInterval(Year::days);
/// @endcode
template<int i, int base = 1>
  struct Timespan
{
  enum {
    seconds = i * base,
    minutes = seconds / 60,
    hours = seconds / 3600,
    days = seconds / 86400
  };
};

/// @brief A Minute.
typedef Timespan<60>
  Minute;

/// @brief An Hour.
typedef Timespan<60, Minute::seconds>
  Hour;

/// @brief A Day.
typedef Timespan<24, Hour::seconds>
  Day;

/// @brief A Week.
typedef Timespan<7, Day::seconds>
  Week;

/// @brief A Month with 30 days.
typedef Timespan<30, Day::seconds>
  ShortMonth;

/// @brief A Month with 31 days.
typedef Timespan<31, Day::seconds>
  LongMonth;

/// @brief A Year with 365 days.
typedef Timespan<365, Day::seconds>
  Year;

/// @brief 64bit timestamp.
typedef ULONGLONG time_64;

/// @brief 32bit timestamp.
typedef time_t time_32;

/// @brief `double` timestamp.
typedef double time_d;

/// @brief Difference between Win32-epoch and unix-epoch in 100-nanoseconds.
static const time_64 EPOCH_DIFF = 116444736000000000LL;

/// @brief Factor seconds to milliseconds.
static const ULONG RATE_DIFF_S_JS = 1000;

/// @brief Factor seconds to 100 nsecs.
static const ULONG RATE_DIFF_UNIX = 10000000;

/// @brief Factor milliseconds to 100 nsecs.
static const ULONG RATE_DIFF_JS = 10000;

/// @brief Converts a `FILETIME` to a `time_64`.
inline time_64 FILETIME2time_64(const FILETIME & aFileTime) {
  return (((time_64)aFileTime.dwLowDateTime) | (((time_64)aFileTime.dwHighDateTime) << 32));
}

/// @brief Win32-timestamps: 64bit, 100 nsecs.
namespace Win32 {
  typedef time_64 Time;
}

/// @brief Unix-timestamps: 32bit, seconds.
namespace Unix {
  typedef time_32 Time;
}

/// @brief JScript-timestamps: `double`, 100 nsecs.
namespace JScript {
  typedef time_d Time;
}

// Win32 timestamps
namespace Win32 {

  /// @brief Unix-timestamp => Win32-timestamp.
  inline Time from(const Unix::Time & aOther)
  {
    return (Time)aOther * RATE_DIFF_UNIX + EPOCH_DIFF;
  }

  /// @brief JScript-timestamp => Win32-timestamp.
  inline Time from(const JScript::Time & aOther)
  {
    return (Time)aOther * RATE_DIFF_JS + EPOCH_DIFF;
  }

  /// @brief Current timestamp as a `Win32::Time`.
  inline Time now()
  {
    FILETIME ft = {0};
    GetSystemTimeAsFileTime(&ft);
    return FILETIME2time_64(ft);
  }
}

// Unix timestamps
namespace Unix {

  /// @brief Win32-timestamp => Unix-timestamp.
  inline Time from(const Win32::Time & aOther)
  {
    return (Time)(aOther - EPOCH_DIFF) / RATE_DIFF_UNIX;
  }

  /// @brief JScript-timestamp => Unix-timestamp.
  inline Time from(const JScript::Time & aOther)
  {
    return (Time)(aOther / RATE_DIFF_S_JS);
  }

  /// @brief Current timestamp as a `Unix::Time`.
  inline Time now()
  {
    return from(Win32::now());
  }
}

// JScript timestamps
namespace JScript {

  // note that we are flooring incoming integer values.
  /// @brief Win32-timestamp => JScript-timestamp.
  inline Time from(const Win32::Time & aOther)
  {
    return (Time)(((Time)aOther - EPOCH_DIFF) / RATE_DIFF_JS);
  }

  /// @brief Unix-timestamp => JScript-timestamp.
  inline Time from(const Unix::Time & aOther)
  {
    return (Time)((Time)aOther * RATE_DIFF_S_JS);
  }

  /// @brief Current timestamp as a `JScript::Time`.
  inline Time now()
  {
    return from(Win32::now());
  }
}

} // namespace Time
} // namespace Ajvar
