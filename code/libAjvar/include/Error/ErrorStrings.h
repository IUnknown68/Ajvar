/**************************************************************************//**
 * @file
 * @brief     Human readable error strings
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/

#pragma once

namespace Ajvar {

/// @brief The `Error` namespaces contains a bunch of functions to get
/// the human readable error string for most of the standard windows errors.
/// It contains the groups Win32 errors (`LRESULT`) and COM errors (`HRESULT`).
namespace Error {

/// @brief Returns the facility of an `HRESULT` error code.
/// Possible values: @see `DErrFACILITY()`
LPCTSTR DErrFACILITY(HRESULT aError);

/// @brief Returns an `LRESULT` error string.
LPCTSTR DErrWin32(LRESULT aError);

/// @brief Returns the severity of an `HRESULT` error code.
/// Possible values are:
/// - `SEVERITY_SUCCESS`
/// - `SEVERITY_ERROR`
LPCTSTR DErrSEVERITY(HRESULT aError);

/// @brief Returns an `HRESULT` error string.
LPCTSTR DErrHRESULT(HRESULT aError);

/// @brief Typedef for a function taking an error code (`DWORD_PTR`) and returning
/// a const TCHAR*
typedef LPCTSTR (*DErr_type)(DWORD_PTR);

/// @{
/// @brief These are the handlers for unknown error codes. By default they point
/// to `_DErr_Default()`, which returns "???".
///
/// You can replace these pointers here by your own implementation to handle
/// your own error codes.
extern DErr_type DErrFACILITYDefault;
extern DErr_type DErrWin32Default;
extern DErr_type DErrSEVERITYDefault;
extern DErr_type DErrHRESULTDefault;
/// @}

} // namespace Error
} // namespace Ajvar
