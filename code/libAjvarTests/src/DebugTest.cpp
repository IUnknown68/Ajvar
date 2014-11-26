#define _DEBUG_COM_OBJECTS
#include "stdafx_tests.h"
#include "AtlModule.h"

//============================================================================
// TESTS
//============================================================================

namespace Ajvar {
namespace Debug {

_D_COMObjectEntryMap _the_D_COMObjectEntryMap;

class ATL_NO_VTABLE DebugTestClass :
  public AjvarComObjectRootEx<DebugTestClass, ATL::CComSingleThreadModel>
{
public:
  typedef ATL::CComObject<DebugTestClass> _ComObject;

  DECLARE_NO_REGISTRY()
  DECLARE_NOT_AGGREGATABLE(DebugTestClass)
  DECLARE_PROTECT_FINAL_CONSTRUCT()

  BEGIN_COM_MAP(DebugTestClass)
  END_COM_MAP()
};


//----------------------------------------------------------------------------
TEST(DebugTest, Leak)
{
  _the_D_COMObjectEntryMap.clear();
  do {
    DebugTestClass::_ComObject * instance = nullptr;
    HRESULT hr = DebugTestClass::_ComObject::CreateInstance(&instance);
    ASSERT_HRESULT_SUCCEEDED(hr);
    instance->AddRef();
  } while(0);
#ifdef _DEBUG
  ASSERT_FALSE(_D_IsObjectMapEmpty());
#else
// release versions never track!
  ASSERT_TRUE(_D_IsObjectMapEmpty());
#endif // def _DEBUG
}


//----------------------------------------------------------------------------
TEST(DebugTest, NoLeak)
{
  _the_D_COMObjectEntryMap.clear();
  do {
    DebugTestClass::_ComObject * instance = nullptr;
    HRESULT hr = DebugTestClass::_ComObject::CreateInstance(&instance);
    ASSERT_HRESULT_SUCCEEDED(hr);
    instance->AddRef();
    instance->Release();
  } while(0);
  ASSERT_TRUE(_D_IsObjectMapEmpty());
}


} //namespace Debug
} //namespace Ajvar
