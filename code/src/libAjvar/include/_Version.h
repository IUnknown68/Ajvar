/**************************************************************************//**
@file
@brief     Main include file for using Ajvar
@author    Arne Seib <arne@salsitasoft.com>
@copyright 2014 Salsita Software (http://www.salsitasoft.com).
***************************************************************************/
#pragma once

namespace Ajvar {

struct _VersionInfo
{
  const ULONG Major;
  const ULONG Minor;
  const ULONG Patch;
  LPCWSTR pre;
  LPCWSTR meta;
};

extern const _VersionInfo Version;


#ifdef CURRENT_VERSION_PARSER
class CurrentVersion
{
public:
  static CurrentVersion & CurrentVersion::Get()
  {
#define _L___(s) L ## s
#define _L__(s) _L___(s)
    static CurrentVersion curVersion(_L__(AJVAR_VERSION));
#undef _L__
#undef _L___
    return curVersion;
  }

  operator LPCWSTR ()
    { return mFullVersion; }

  ULONG Major()
    { return mMajor; }

  ULONG Minor()
    { return mMinor; }

  ULONG Patch()
    { return mPatch; }

  LPCWSTR Pre()
    { return mPre; }

  LPCWSTR Meta()
    { return mMeta; }

private:
  // not implemented:
  CurrentVersion();
  CurrentVersion(const CurrentVersion &);

  CurrentVersion(LPCWSTR fullVersion) :
    mFullVersion(nullptr), mMajor(0), mMinor(0), mPatch(0), mPre(nullptr), mMeta(nullptr)
  {
    static LPCWSTR empty = L"";
    size_t len = std::wcslen(fullVersion) + 1;
    mFullVersion = new wchar_t[len];
    wcscpy_s(mFullVersion, len, fullVersion);
    wchar_t * p = mFullVersion;

    // major
    mMajor = std::wcstoul(p, &p, 10);
    if (!(*p)) {
      return;
    }
    ++p;
    // minor
    mMinor = std::wcstoul(p, &p, 10);
    if (!(*p)) {
      return;
    }
    ++p;
    // patch
    mPatch = std::wcstoul(p, &p, 10);
    if (!(*p)) {
      return;
    }

    // pre
    wchar_t * context = nullptr;
    mPre = wcstok_s(p, L"-+", &context);
    if (mPre) {
      // meta
      mMeta = wcstok_s(nullptr, L"-+", &context);
      if (!mMeta) {
        mMeta = empty;
      }
    }
    else {
      mPre = empty;
    }
  }

  ~CurrentVersion()
  {
    if (nullptr != mFullVersion) {
      delete [] mFullVersion;
      mFullVersion = nullptr;
    }
  }

  LPWSTR mFullVersion;
  ULONG mMajor;
  ULONG mMinor;
  ULONG mPatch;
  LPCWSTR mPre;
  LPCWSTR mMeta;
};

inline CurrentVersion & Version()
{
  return CurrentVersion::Get();
}
#endif // def 
} // namespace Ajvar
