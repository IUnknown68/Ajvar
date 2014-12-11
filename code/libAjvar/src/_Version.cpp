/**************************************************************************//**
 * @file
 * @brief     Main include file for `Ajvar::Dispatch::Ex`
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/
#pragma once

#include "stdafx.h"
#include "_Version.h"

namespace Ajvar {

#define _L___(s) L ## s
#define _L__(s) _L___(s)
const _VersionInfo Version =
{
  AJ_VER_MAJ,
  AJ_VER_MIN,
  AJ_VER_PAT,
  _L__(AJ_VER_PRE),
  _L__(AJ_VER_MET)
};
#undef _L__
#undef _L___

} // namespace Ajvar
