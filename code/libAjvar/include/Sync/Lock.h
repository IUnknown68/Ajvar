/**************************************************************************//**
@file
@brief     class Ajvar::Sync::Lock
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {
namespace Sync {

/// @brief Scoped lock for `CRITICAL_SECTION`.
/// @details `Lock` is a simple scoped lock class operating on a
/// `CRITICAL_SECTION`. It locks on construction and unlocks on destruction.
class Lock
{
public:
  /// @brief Constructor: Locks the `CRITICAL_SECTION` via `EnterCriticalSection()`.
  Lock(CRITICAL_SECTION & cs)
    : mCS(cs)
  {
    ::EnterCriticalSection(&mCS);
  }

  /// @brief Destructor: Unlocks the `CRITICAL_SECTION` via `LeaveCriticalSection()`.
  ~Lock()
  {
    ::LeaveCriticalSection(&mCS);
  }

private:
  NO_COPY(Lock)
  /// @brief Reference to the `CRITICAL_SECTION` to work on.
  CRITICAL_SECTION & mCS;
};

} // namespace Sync
} // namespace Ajvar
