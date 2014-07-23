/**************************************************************************//**
@file
@brief     Global inteface table handling.
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/
#pragma once

namespace Ajvar {

#ifdef AJ_USE_GLOBALINTERFACE

/// @brief A class for storing COM objects in the global interface table.
/// @details Contains a `IGlobalInterfaceTable` and a cookie for an object
/// stored in the global interface table. Can set, get and revoke the object.
template <typename TInterface>
  class GlobalInterface
{
public:
  /// @brief Empty constructor.
  GlobalInterface() :
    mCookie(0)
  {
  }

  /// @brief Constructor accepting an object pointer.
  GlobalInterface(TInterface * aObject) :
    mCookie(0)
  {
    Set(aObject);
  }

  /// @brief Destructor.
  ~GlobalInterface()
  {
    Unset();
  }

  /// @brief Sets an object in the global interface table. Fails if the
  /// table can't be accessed or we are already guiding another object.
  HRESULT Set(TInterface * aObject)
  {
    if (0 != mCookie) {
      return E_UNEXPECTED;
    }
    HRESULT hr = E_FAIL;
    ATL::CComPtr<IGlobalInterfaceTable> table = GetTable(&hr);
    if (!table) {
      return hr;
    }
    return table->RegisterInterfaceInGlobal(static_cast<IUnknown*>(aObject), __uuidof(TInterface), &mCookie);
  }

  /// @brief Removes our object from the global interface table.
  HRESULT Unset()
  {
    if (0 == mCookie) {
      return E_UNEXPECTED;
    }
    HRESULT hr = E_FAIL;
    ATL::CComPtr<IGlobalInterfaceTable> table = GetTable(&hr);
    if (!table) {
      return hr;
    }
    hr = table->RevokeInterfaceFromGlobal(mCookie);
    mCookie = 0;
    return hr;
  }

  /// @brief Gets the object from the global interface table.
  ATL::CComPtr<TInterface> Get(HRESULT * aHrPtr = nullptr)
  {
    ATL::CComPtr<TInterface> object;
    HRESULT hr = E_UNEXPECTED;
    ATL::CComPtr<IGlobalInterfaceTable> table = GetTable(&hr);
    if (table && (mCookie != 0)) {
      hr = table->GetInterfaceFromGlobal(mCookie, __uuidof(TInterface), (void**)&object.p);
    }
    if (nullptr != aHrPtr) {
      (*aHrPtr) = hr;
    }
    return object;
  }

  /// @brief Checks if we currently hold an object.
  bool IsEmpty()
  {
    return (0 == mCookie);
  }

private:
  NO_COPY(GlobalInterface)

  /// @brief Get the global interface table.
  ATL::CComPtr<IGlobalInterfaceTable> GetTable(HRESULT * aHrPtr = nullptr)
  {
    if (!mGlobalInterfaceTable) {
      HRESULT hr = mGlobalInterfaceTable.CoCreateInstance(CLSID_StdGlobalInterfaceTable);
      if (nullptr != aHrPtr) {
        (*aHrPtr) = hr;
      }
    }
    return mGlobalInterfaceTable;
  }

  ATL::CComPtr<IGlobalInterfaceTable> mGlobalInterfaceTable;
  DWORD mCookie;
};

#endif // def AJ_USE_GLOBALINTERFACE

} // namespace Ajvar
