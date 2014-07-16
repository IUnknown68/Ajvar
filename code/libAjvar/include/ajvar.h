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

#include "Dispatch/Connector.h"
#include "Dispatch/RefVariant.h"
#include "Dispatch/Object.h"
#include "Dispatch/Dispatch.h"
#include "Dispatch/Ex/Ex.h"
#include "Dispatch/Ex/Properties.h"
#include "Ajvartime.h"

