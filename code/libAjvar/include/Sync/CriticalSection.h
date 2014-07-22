/**************************************************************************//**
@file
@brief     class Ajvar::Sync::CriticalSection
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/
#pragma once

namespace Ajvar {
namespace Sync {

/// @brief `CriticalSection` is just a wrapper around `CRITICAL_SECTION`.
class CriticalSection :
  public CRITICAL_SECTION
{
public:
  /// @brief Constructor: Initializes the `CRITICAL_SECTION`.
  CriticalSection()
  {
    ::InitializeCriticalSection(static_cast<CRITICAL_SECTION*>(this));
  }

  /// @brief Destructor: Deletes the `CRITICAL_SECTION`.
  virtual ~CriticalSection()
  {
    ::DeleteCriticalSection(static_cast<CRITICAL_SECTION*>(this));
  }

private:
  NO_COPY(CriticalSection)
};

} // namespace Sync
} // namespace Ajvar
