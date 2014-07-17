#include <stdafx.h>
#include <include/ajvar.h>

//============================================================================
// TESTS
//============================================================================

#define ONE_SEC (1000*1000*10)
#define EPOCH_DIFF     116444736000000000LL

namespace Ajvar {
namespace Time {

//----------------------------------------------------------------------------
// Compare "now()" values
TEST(AnchoTimeTest, Now)
{
  auto tWin32 = Win32::now();
  auto tUnix = Unix::now();
  auto tJScript = JScript::now();

  ULONGLONG tUnixWin32 = (((ULONGLONG)tUnix * ONE_SEC) + 116444736000000000LL) - tWin32;
  ASSERT_LT(_abs64(tUnixWin32), ONE_SEC);

  ULONGLONG tJScriptWin32 = (((ULONGLONG)tJScript * 10000) + 116444736000000000LL) - tWin32;
  ASSERT_LT(_abs64(tJScriptWin32), ONE_SEC);
}

//----------------------------------------------------------------------------
// Import Win32 -> Unix
TEST(AnchoTimeTest, ImportWin32IntoUnix)
{
  auto tWin32 = Win32::now();
  auto tUnix = Unix::now();

  auto imported = Unix::from(tWin32);
  ASSERT_LT(_abs64(tUnix - imported), 1);
}

//----------------------------------------------------------------------------
// Import JScript -> Unix
TEST(AnchoTimeTest, ImportJScriptIntoUnix)
{
  auto tJScript = JScript::now();
  auto tUnix = Unix::now();

  auto imported = Unix::from(tJScript);
  ASSERT_LT(_abs64(tUnix - imported), 1);
}

//----------------------------------------------------------------------------
// Import Win32 -> JScript
TEST(AnchoTimeTest, ImportWin32IntoJScript)
{
  auto tJScript = JScript::now();
  auto tWin32 = Win32::now();

  auto imported = JScript::from(tWin32);
  ASSERT_LT(abs(tJScript - imported), 1000);
}

//----------------------------------------------------------------------------
// Import Unix -> JScript
TEST(AnchoTimeTest, ImportUnixIntoJScript)
{
  auto tJScript = JScript::now();
  auto tUnix = Unix::now();

  auto imported = JScript::from(tUnix);
  ASSERT_LT(abs(tJScript - imported), 1000);
}

//----------------------------------------------------------------------------
// Import JScript -> Win32
TEST(AnchoTimeTest, ImportJScriptIntoWin32)
{
  auto tWin32 = Win32::now();
  auto tJScript = JScript::now();

  auto imported = Win32::from(tJScript);
  ASSERT_LT(_abs64(tWin32 - imported), ONE_SEC);
}

//----------------------------------------------------------------------------
// Import Unix -> Win32
TEST(AnchoTimeTest, ImportUnixIntoWin32)
{
  auto tWin32 = Win32::now();
  auto tUnix = Unix::now();

  auto imported = Win32::from(tUnix);
  ASSERT_LT(_abs64(tWin32 - imported), ONE_SEC);
}

} //namespace Time
} //namespace Ajvar
