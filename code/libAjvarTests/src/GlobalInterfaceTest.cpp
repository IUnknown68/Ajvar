#include <stdafx.h>
#include <include/ajvar.h>
#include "AtlModule.h"

//============================================================================
// TESTS
//============================================================================

namespace Ajvar {

#ifdef AJ_USE_GLOBALINTERFACE

MIDL_INTERFACE("D72A6CA5-6691-4DC5-A6F4-76F15D6ADFBF")
IGlobalInterfaceTestClassInterface : public IUnknown
{
public:
};

GlobalInterface<IGlobalInterfaceTestClassInterface> global;
IGlobalInterfaceTestClassInterface * unkFromGlobalInterface = nullptr;

class ATL_NO_VTABLE GlobalInterfaceTestClass :
  public ATL::CComObjectRootEx<ATL::CComSingleThreadModel>,
  public IGlobalInterfaceTestClassInterface
{
public:
  typedef ATL::CComObject<GlobalInterfaceTestClass> _ComObject;

  DECLARE_PROTECT_FINAL_CONSTRUCT()

  BEGIN_COM_MAP(GlobalInterfaceTestClass)
    COM_INTERFACE_ENTRY(IGlobalInterfaceTestClassInterface)
  END_COM_MAP()
};

DWORD WINAPI ThreadProc(LPVOID lpParameter)
{
  ::CoInitializeEx(NULL, COINIT_MULTITHREADED);
  ATL::CComPtr<IGlobalInterfaceTestClassInterface> unk = global.Get();
  unkFromGlobalInterface = unk.p;
  ::CoUninitialize();
  return 0;
}

//----------------------------------------------------------------------------
TEST(GlobalInterfaceTest, Get)
{
  ::CoInitializeEx(NULL, COINIT_MULTITHREADED);
  GlobalInterfaceTestClass::_ComObject * instance = nullptr;
  HRESULT hr = GlobalInterfaceTestClass::_ComObject::CreateInstance(&instance);
  ASSERT_HRESULT_SUCCEEDED(hr);
  ATL::CComPtr<IGlobalInterfaceTestClassInterface> unk(instance);

  global.Set(unk);
  HANDLE h = ::CreateThread(nullptr, 0, ThreadProc, nullptr, 0, nullptr);
  WaitForSingleObject(h, 1000);
  ASSERT_NE(nullptr, unkFromGlobalInterface);
  ::CoUninitialize();
}

#endif // def AJ_USE_GLOBALINTERFACE

} //namespace Ajvar
