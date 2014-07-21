/**************************************************************************//**
@file
@brief     class Ajvar::Sync::Lock
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/
#pragma once

namespace Ajvar {
namespace Sync {

/// @brief `Lock` is a simple scoped lock class operating on a
/// `CRITICAL_SECTION`. It locks on construction and unlocks on destruction.
class Lock
{
public:
  /// @brief Constructor: Enters (locks) the `CRITICAL_SECTION`.
  Lock(CRITICAL_SECTION & cs)
    : mCS(cs)
  {
    ::EnterCriticalSection(&mCS);
  }

  /// @brief Destructor: Leaves (unlocks) the `CRITICAL_SECTION`.
  ~Lock()
  {
    ::LeaveCriticalSection(&mCS);
  }

private:
  NO_COPY(Lock)
  CRITICAL_SECTION & mCS;
};

} // namespace Sync
} // namespace Ajvar
