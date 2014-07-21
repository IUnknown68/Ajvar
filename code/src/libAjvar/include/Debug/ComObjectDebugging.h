/**************************************************************************//**
@file
@brief     class Ajvar::Debug::AjvarComObjectRootEx and related stuff.
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/

#pragma once

namespace Ajvar {

  
/// @brief Helpers for debugging COM
/// @details `Ajvar::Debug` contains some classes for debugging COM object
/// lifetime management. There is a global map storing information about
/// all COM objects implemented using `AjvarComObjectRootEx`, which replaces
/// `CComObjectRootEx`. When a new object is created it gets added to the map,
/// when it gets destroyed, it gets removed. So to see which objects are
/// leaking you would dump the map right before your process exits.
/// @note This is not thread safe, and the implementation is rather simple.
/// It was originally created for one single bug, so use with care!
/// @note Uses RTTI.
///
/// Using it is quite simple:
/// - Define `_DEBUG_COM_OBJECTS` in your debug builds before including
/// Ajvar.h.
/// - Derive your ATL COM objects from `AjvarComObjectRootEx` instead of
/// `CComObjectRootEx` with the additional first argument of your implementing
/// class:
/// @code{.cpp}
/// class ATL_NO_VTABLE CMyComObject :
///   public AjvarComObjectRootEx<CMyComObject, CComSingleThreadModel>,
/// ...
/// @endcode
/// - Before your process exits, after all objects are supposed to be
/// destroyed, call `_D_IsObjectMapEmpty()` to see if you have leaking
/// objects, and optionally `_D_DumpCOMObjects()` to list them in the output
/// window.
namespace Debug {

#ifdef _DEBUG_COM_OBJECTS

// -------------------------------------------------------------------------
/// @brief Class for a key in the global COM object map. The key contains all
/// the information, while the value contains just a weak pointer.
struct _D_COMObjectKey
{
  /// @brief Full type name - not used anywhere.
  std::string typeName;

  /// @brief Short type name - used as part of the key.
  std::string typeNameShort;

  /// @brief Raw pointer to the object.
  ATL::CComObjectRootBase * ptr;

  /// @brief Constructor.
  _D_COMObjectKey(const char * t, const char * ts, ATL::CComObjectRootBase * const& p) :
    typeName(t), typeNameShort(ts), ptr(p)
  {
  }

  /// @brief Copy constructor.
  _D_COMObjectKey(const _D_COMObjectKey & aOther) :
    typeName(aOther.typeName), typeNameShort(aOther.typeNameShort), ptr(aOther.ptr)
  {
  }

  /// @brief Assignment.
  _D_COMObjectKey & operator = (const _D_COMObjectKey & aOther)
  {
    typeName = aOther.typeName;
    typeNameShort = aOther.typeNameShort;
    ptr = aOther.ptr;
  }

  /// @brief Prints info about the object.
  void dump() const
  {
    ATLTRACE("%s at 0x%08x REF %i\n", typeNameShort.c_str(), ptr, (ptr) ? ptr->m_dwRef : -1);
  }
};

// -------------------------------------------------------------------------
/// @brief Key hash for storing key in map.
struct _D_COMObjectKeyHash {
 std::size_t operator()(const _D_COMObjectKey& k) const
 {
     return std::hash<std::string>()(k.typeName) ^
            (std::hash<ULONG>()(reinterpret_cast<ULONG>(k.ptr)) << 1);
 }
};
 
// -------------------------------------------------------------------------
/// @brief Key equal for comparision.
struct _D_COMObjectKeyEqual {
 bool operator()(const _D_COMObjectKey& lhs, const _D_COMObjectKey& rhs) const
 {
    return (lhs.typeName == rhs.typeName) && (lhs.ptr == rhs.ptr);
 }
};

// -------------------------------------------------------------------------
/// @brief Global map for storing all `AjvarComObjectBase` instances.
typedef std::unordered_map<
  _D_COMObjectKey,
  ATL::CComObjectRootBase*,
  _D_COMObjectKeyHash,
  _D_COMObjectKeyEqual> _D_COMObjectEntryMap;

/// @note The actual map must be in one of your cpp files.
extern _D_COMObjectEntryMap _the_D_COMObjectEntryMap;

// -------------------------------------------------------------------------
/// @brief Baseclass replacing `CComObjectRootBase`.
template<class Impl> class AjvarComObjectBase :
  public ATL::CComObjectRootBase
{
public:
  typedef Impl _Impl;

protected:
  /// @brief Called after an instance was created. Adds the object to the global map.
  void _D_COM_Created()
  {
    _D_COMObjectKey key(typeid(*static_cast<_Impl*>(this)).name(), typeid(_Impl).name(), this);
    _the_D_COMObjectEntryMap[key] = this;
    ATLTRACE(_T("+++ (%i) NEW: "), _the_D_COMObjectEntryMap.size());
    key.dump();
  }

  /// @brief Called from `AddRef`.
  LONG _D_COM_AddRef(LONG nRef)
  {
    _D_COMObjectKey key(typeid(*static_cast<_Impl*>(this)).name(), typeid(_Impl).name(), this);
    ATLTRACE(_T("+ AddRef(%i): "), nRef);
    key.dump();
    return nRef;
  }

  /// @brief Called from `Release`.
  LONG _D_COM_Release(long nRef)
  {
    _D_COMObjectKey key(typeid(*static_cast<_Impl*>(this)).name(), typeid(_Impl).name(), this);
    ATLTRACE(_T("+ Release(%i): "), nRef);
    key.dump();
    if (0 == nRef) {
      auto it = _the_D_COMObjectEntryMap.find(_D_COMObjectKey(typeid(*static_cast<_Impl*>(this)).name(), typeid(_Impl).name(), this));
      if (it != _the_D_COMObjectEntryMap.end()) {
        ATLTRACE(_T("--- (%i) DELETED: "), _the_D_COMObjectEntryMap.size()-1);
        it->first.dump();
        _the_D_COMObjectEntryMap.erase(it);
      }
    }
    return nRef;
  }
};

// -------------------------------------------------------------------------
/// @brief Replacement for `CComObjectRootEx`. Mostly the same like
/// `CComObjectRootEx`, except that the lifetime management methods notify
/// the global map.
template <class Impl, class ThreadModel>
class AjvarComObjectRootEx : 
  public AjvarComObjectBase<Impl>
{
public:
  /// @{
  /// same as `CComObjectRootEx`
  typedef ThreadModel _ThreadModel;
  typedef typename _ThreadModel::AutoCriticalSection _CritSec;
  typedef typename _ThreadModel::AutoDeleteCriticalSection _AutoDelCritSec;
  typedef ATL::CComObjectLockT<_ThreadModel> ObjectLock;
  /// @}

  ~AjvarComObjectRootEx() 
  {
  }

  /// @brief Equivalent for `CComObjectRootEx::InternalAddRef()`
  ULONG InternalAddRef(BOOL bInternal = FALSE)
  {
    ATLASSUME(m_dwRef != -1L);
#ifdef _DEBUG
    return (bInternal) ? _ThreadModel::Increment(&m_dwRef) : _D_COM_AddRef(_ThreadModel::Increment(&m_dwRef));
#else
    return _ThreadModel::Increment(&m_dwRef);
#endif  
  }

  /// @brief Equivalent for `CComObjectRootEx::InternalRelease()`
  ULONG InternalRelease(BOOL bInternal = FALSE)
  {
#ifdef _DEBUG
    LONG nRef = _ThreadModel::Decrement(&m_dwRef);
    if (nRef < -(LONG_MAX / 2)) {
      ATLASSERT(0 && _T("Release called on a pointer that has already been released"));
    }
    return (bInternal) ? nRef : _D_COM_Release(nRef);
#else
    return _ThreadModel::Decrement(&m_dwRef);
#endif
  }

  /// @brief Equivalent for `CComObjectRootEx::_AtlInitialConstruct()`
  HRESULT _AtlInitialConstruct()
  {
    _D_COM_Created();
    return m_critsec.Init();
  }

  /// @brief Equivalent for `CComObjectRootEx::Lock()`
  void Lock() 
  {
    m_critsec.Lock();
  }

  /// @brief Equivalent for `CComObjectRootEx::Unlock()`
  void Unlock() 
  {
    m_critsec.Unlock();
  }
private:

  /// @brief Equivalent for `CComObjectRootEx::m_critsec`
  _AutoDelCritSec m_critsec;
};

// -------------------------------------------------------------------------
/// @brief Specialization for `CComSingleThreadModel`.
template <class Impl>
class AjvarComObjectRootEx<Impl, ATL::CComSingleThreadModel> : 
  public AjvarComObjectBase<Impl>
{
public:
  /// @{
  /// same as `CComObjectRootEx`
  typedef ATL::CComSingleThreadModel _ThreadModel;
  typedef _ThreadModel::AutoCriticalSection _CritSec;
  typedef _ThreadModel::AutoDeleteCriticalSection _AutoDelCritSec;
  typedef ATL::CComObjectLockT<_ThreadModel> ObjectLock;
  /// @}

  ~AjvarComObjectRootEx() {}

  /// @brief Equivalent for `CComObjectRootEx::InternalAddRef()`
  ULONG InternalAddRef(BOOL bInternal = FALSE)
  {
    ATLASSUME(m_dwRef != -1L);
#ifdef _DEBUG
    return (bInternal) ? _ThreadModel::Increment(&m_dwRef) : _D_COM_AddRef(_ThreadModel::Increment(&m_dwRef));
#else
    return _ThreadModel::Increment(&m_dwRef);
#endif  
  }

  /// @brief Equivalent for `CComObjectRootEx::InternalRelease()`
  ULONG InternalRelease(BOOL bInternal = FALSE)
  {
#ifdef _DEBUG
    long nRef = _ThreadModel::Decrement(&m_dwRef);
    if (nRef < -(LONG_MAX / 2)) {
      ATLASSERT(0 && _T("Release called on a pointer that has already been released"));
    }
    return (bInternal) ? nRef : _D_COM_Release(nRef);
#else
    return _ThreadModel::Decrement(&m_dwRef);
#endif  
  }

  /// @brief Equivalent for `CComObjectRootEx::_AtlInitialConstruct()`
  HRESULT _AtlInitialConstruct()
  {
    _D_COM_Created();
    return S_OK;
  }

  /// @brief Equivalent for `CComObjectRootEx::Lock()`
  void Lock() 
  {
  }

  /// @brief Equivalent for `CComObjectRootEx::Unlock()`
  void Unlock() 
  {
  }
};

/// @brief Replacement for `DECLARE_PROTECT_FINAL_CONSTRUCT`, includes an argument
/// for InternalAddRef() and InternalRelease() to signalize internal
/// calls.
#undef DECLARE_PROTECT_FINAL_CONSTRUCT
#define DECLARE_PROTECT_FINAL_CONSTRUCT()\
  void InternalFinalConstructAddRef() {\
    InternalAddRef(true);\
  }\
  void InternalFinalConstructRelease() {\
    InternalRelease(true);\
  }

// -------------------------------------------------------------------------
/// @brief Checks if _the_D_COMObjectEntryMap is empty. Quick way to check
/// if there are leaking objects.
inline bool _D_IsObjectMapEmpty()
{
  return _the_D_COMObjectEntryMap.empty();
}

// -------------------------------------------------------------------------
/// @brief Prints the contents of the global map.
// _D_DumpCOMObjects()
inline void _D_DumpCOMObjects()
{
  ATLTRACE(_T("************************************************************************************\n"));
  if (!_D_IsObjectMapEmpty()) {
    ATLTRACE(_T("* DUMPING COM OBJECTS: %i\n"), _the_D_COMObjectEntryMap.size());
    for (auto it = _the_D_COMObjectEntryMap.begin(); it != _the_D_COMObjectEntryMap.end(); ++it) {
      ATLTRACE(_T("* "));
      it->first.dump();
    }
  }
  else {
    ATLTRACE(_T("* no COM objects\n"));
  }
  ATLTRACE(_T("************************************************************************************\n"));
}

// =========================================================================
#else // def _DEBUG_COM_OBJECTS
// release versions of debugging stuff
// =========================================================================

inline void _D_DumpCOMObjects()
{
}

inline bool _D_IsObjectMapEmpty()
{
  return true;
}

template <class Impl, class ThreadModel>
  class AjvarComObjectRootEx :
    public ATL::CComObjectRootEx<ThreadModel>
{
};

// =========================================================================
#endif // def _DEBUG_COM_OBJECTS
// =========================================================================

} // namespace Debug
} // namespace Ajvar
