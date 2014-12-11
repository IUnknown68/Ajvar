#include <stdafx.h>
#include <include/ajvar.h>

struct TestDataTuple
{
  DISPID id;
  LPCOLESTR name;
  ATL::CComVariant value;
};

using namespace ATL;

//============================================================================
// Fixture
//============================================================================

// actual testdata, 3 entries
static TestDataTuple testData[] = {
  {1, L"foo", L"foovalue"},
  {2, L"bar", L"barvalue"},
  {3, L"bas", L"basvalue"}
};

//============================================================================
// class PropertiesTest

class PropertiesTest :
  public ::testing::Test,
  public ::Ajvar::Dispatch::Ex::Properties
{
protected:
  // additional setup helpers
  void setPropertyNames(int index = -1);
  void setPropertyValues(int index = -1);
  void assureRemoved(int index);
  void verify(int index = -1);
};

//----------------------------------------------------------------------------
// setup: add 3 property names, no values
void PropertiesTest::setPropertyNames(int index)
{
  if (index < 0) {
    for (int n = 0; n < _countof(testData); n++) {
      setPropertyNames(n);
    }
    return;
  }
  if(index >= _countof(testData)) {
    throw std::range_error("Invalid index for testdata");
  }
  DISPID did = DISPID_UNKNOWN;
  ASSERT_HRESULT_SUCCEEDED(getDispID(testData[index].name, &did, TRUE));
  ASSERT_EQ(testData[index].id, did);
}

//----------------------------------------------------------------------------
// setup: add property name and value
void PropertiesTest::setPropertyValues(int index)
{
  if (index < 0) {
    for (int n = 0; n < _countof(testData); n++) {
      setPropertyValues(n);
    }
    return;
  }
  if(index >= _countof(testData)) {
    throw std::range_error("Invalid index for testdata");
  }
  DISPID did = DISPID_UNKNOWN;
  ASSERT_HRESULT_SUCCEEDED(getDispID(testData[index].name, &did, TRUE));
  ASSERT_EQ(testData[index].id, did);
  ASSERT_HRESULT_SUCCEEDED(putValue(did, testData[index].value));
  verify(index);
}

//----------------------------------------------------------------------------
// assure value got removed: Assumes that originally 3 values were set, one
// got removed
void PropertiesTest::assureRemoved(int index)
{
  if(index < 0 || index >= _countof(testData)) {
    throw std::range_error("Invalid index for testdata");
  }
  // check that property does not exist any more...
  EXPECT_EQ(2, mValues.size());
  // ... but still is a known name
  EXPECT_EQ(3, mNames.size());
  EXPECT_EQ(3, mNameIDs.size());

  // - get name for ID: should still work
  CComBSTR name;
  EXPECT_HRESULT_SUCCEEDED(getName(testData[index].id, &name));
  EXPECT_STREQ(testData[index].name, name.m_str);

  // - get dispid: shoud still work
  DISPID did = DISPID_UNKNOWN;
  EXPECT_HRESULT_SUCCEEDED(getDispID(testData[index].name, &did, FALSE));
  EXPECT_EQ(testData[index].id, did);

  // - get value: should fail
  CComVariant vt;
  ASSERT_EQ(DISP_E_MEMBERNOTFOUND, getValue(testData[index].id, &vt));
  ASSERT_EQ(VT_EMPTY, vt.vt);
}

//----------------------------------------------------------------------------
// verfy a (or all) properties
void PropertiesTest::verify(int index)
{
  if (index < 0) {
    for (int n = 0; n < _countof(testData); n++) {
      verify(n);
    }
    return;
  }
  if(index >= _countof(testData)) {
    throw std::range_error("Invalid index for testdata");
  }

  {
    auto it = mNameIDs.find(testData[index].name);
    ASSERT_TRUE((it != mNameIDs.end()));
    ASSERT_EQ(testData[index].id, it->second);
  }

  {
    auto it = mNames.find(testData[index].id);
    ASSERT_TRUE((it != mNames.end()));
    ASSERT_STREQ(testData[index].name, it->second);
  }

  {
    auto it = mValues.find(testData[index].id);
    ASSERT_TRUE((it != mValues.end()));
    ASSERT_EQ(testData[index].value.vt, it->second.vt);
    ASSERT_STREQ(testData[index].value.bstrVal, it->second.bstrVal);
  }
}

//============================================================================
// TESTS
//============================================================================

//----------------------------------------------------------------------------
// NULL pointer checks
TEST_F(PropertiesTest, NullPointerChecks)
{
  //EXPECT_EQ(E_INVALIDARG, putValue(1, NULL));
  EXPECT_EQ(E_POINTER, getValue(1, NULL));
  EXPECT_EQ(E_POINTER, getName(1, NULL));
  EXPECT_EQ(E_POINTER, getDispID(L"foo", NULL));
  EXPECT_EQ(E_POINTER, enumNextDispID(0, 1, NULL));
}

//----------------------------------------------------------------------------
// test if getNextDISPID works
TEST_F(PropertiesTest, getNextDISPID)
{
  setPropertyNames();
  EXPECT_EQ(3, mLastDISPID);
  EXPECT_EQ(4, getNextDISPID());
}

//----------------------------------------------------------------------------
// add 3 property names, no values
TEST_F(PropertiesTest, AddNewPropertyName)
{
  setPropertyNames();
}

//----------------------------------------------------------------------------
// get name for a non-existing property
TEST_F(PropertiesTest, GetNonExistingProperty)
{
  DISPID did = DISPID_UNKNOWN;
  EXPECT_EQ(DISP_E_UNKNOWNNAME, getDispID(L"doesnotexist", &did, FALSE));
  EXPECT_EQ(DISPID_UNKNOWN, did);
}

//----------------------------------------------------------------------------
// get name for an existing property without value
TEST_F(PropertiesTest, GetExistingProperty)
{
  setPropertyNames();
  int n = 1;

  DISPID did = DISPID_UNKNOWN;
  EXPECT_HRESULT_SUCCEEDED(getDispID(testData[n].name, &did, FALSE));
  EXPECT_EQ(testData[n].id, did);
}

//----------------------------------------------------------------------------
// add a property with it's value
TEST_F(PropertiesTest, AddNewPropertyValue)
{
  setPropertyNames();
  setPropertyValues(1);
}

//----------------------------------------------------------------------------
// get value for a property that exist, but has no value
TEST_F(PropertiesTest, GetNonExistingValue)
{
  setPropertyNames();
  int n = 1;

  CComVariant vt;
  ASSERT_EQ(DISP_E_MEMBERNOTFOUND, getValue(testData[n].id, &vt));
  ASSERT_EQ(VT_EMPTY, vt.vt);
}

//----------------------------------------------------------------------------
// get property value that exists
TEST_F(PropertiesTest, GetExistingPropertyValue)
{
  setPropertyNames();
  int n = 1;
  setPropertyValues(n);

  DISPID did = DISPID_UNKNOWN;
  ASSERT_HRESULT_SUCCEEDED(getDispID(testData[n].name, &did, FALSE));
  ASSERT_EQ(testData[n].id, did);

  CComVariant vt;
  EXPECT_HRESULT_SUCCEEDED(getValue(did, &vt));
  EXPECT_EQ(VT_BSTR, vt.vt);
  EXPECT_STREQ(testData[n].value.bstrVal, vt.bstrVal);
}

//----------------------------------------------------------------------------
// get name from dispid, non-existing
TEST_F(PropertiesTest, GetNonExistingNameFromDISPID)
{
  setPropertyNames();
  CComBSTR name;
  EXPECT_EQ(DISP_E_MEMBERNOTFOUND, getName(123, &name));
  EXPECT_EQ(NULL, name.m_str);

  // also test that out value stays untouched
  name = L"test";
  EXPECT_EQ(DISP_E_MEMBERNOTFOUND, getName(123, &name));
  ASSERT_NE((LPCOLESTR)NULL, name.m_str);
  EXPECT_STREQ(L"test", name.m_str);
}

//----------------------------------------------------------------------------
// get name from dispid, existing
TEST_F(PropertiesTest, GetExistingNameFromDISPID)
{
  setPropertyNames();
  int n = 1;
  setPropertyValues(n);

  CComBSTR name;
  EXPECT_HRESULT_SUCCEEDED(getName(testData[n].id, &name));
  ASSERT_NE((LPCOLESTR)NULL, name.m_str);
  EXPECT_STREQ(testData[n].name, name.m_str);

}

//----------------------------------------------------------------------------
// get name from dispid, existing, but no value
TEST_F(PropertiesTest, GetExistingNameFromDISPIDNoValue)
{
  setPropertyNames();
  int n = 1;

  CComBSTR name;
  EXPECT_HRESULT_SUCCEEDED(getName(testData[n].id, &name));
  ASSERT_NE((LPCOLESTR)NULL, name.m_str);
  EXPECT_STREQ(testData[n].name, name.m_str);

}

//----------------------------------------------------------------------------
// remove property by name, value exists
TEST_F(PropertiesTest, RemovePropertyByName)
{
  setPropertyNames();
  setPropertyValues();
  int n = 1;
  EXPECT_HRESULT_SUCCEEDED(remove(testData[n].name));
  assureRemoved(n);
}

//----------------------------------------------------------------------------
// remove property by name, value does not exist
TEST_F(PropertiesTest, RemovePropertyWithoutValueByName)
{
  setPropertyNames();
  setPropertyValues(1);
  EXPECT_EQ(DISP_E_UNKNOWNNAME, remove(testData[0].name));
}

//----------------------------------------------------------------------------
// remove property by name, property does not exist
TEST_F(PropertiesTest, RemoveNonexistingPropertyByName)
{
  setPropertyNames();
  EXPECT_EQ(DISP_E_UNKNOWNNAME, remove(testData[0].name));
}

//----------------------------------------------------------------------------
// remove property by id, value exists
TEST_F(PropertiesTest, RemovePropertyByDISPID)
{
  setPropertyNames();
  setPropertyValues();
  int n = 1;
  EXPECT_HRESULT_SUCCEEDED(remove(testData[n].id));
  assureRemoved(n);
}

//----------------------------------------------------------------------------
// remove property by id, value does not exist
TEST_F(PropertiesTest, RemovePropertyWithoutValueByDISPID)
{
  setPropertyNames();
  setPropertyValues(1);
  EXPECT_EQ(DISP_E_UNKNOWNNAME, remove(testData[0].id));
}

//----------------------------------------------------------------------------
// remove property by id, property does not exist
TEST_F(PropertiesTest, RemoveNonexistingPropertyByDISPID)
{
  setPropertyNames();
  EXPECT_EQ(DISP_E_UNKNOWNNAME, remove(testData[0].id));
}

//----------------------------------------------------------------------------
// remove property by id, property does not exist
TEST_F(PropertiesTest, PutNonexistingProperty)
{
  setPropertyNames();
  CComVariant(0);
  EXPECT_EQ(DISP_E_MEMBERNOTFOUND, putValue(99999, CComVariant(0)));
}

//----------------------------------------------------------------------------
// iterating tests

//----------------------------------------------------------------------------
// full data, iterate
TEST_F(PropertiesTest, IterateFullProperties)
{
  setPropertyNames();
  setPropertyValues();

  DISPID did = DISPID_STARTENUM;
  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[0].id, did);

  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[1].id, did);

  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[2].id, did);

  ASSERT_EQ(S_FALSE, enumNextDispID(0, did, &did));
}

//----------------------------------------------------------------------------
// partial data, iterate
TEST_F(PropertiesTest, IteratePartialProperties)
{
  setPropertyNames();
  setPropertyValues(0);
  setPropertyValues(2);

  DISPID did = DISPID_STARTENUM;
  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[0].id, did);

  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[2].id, did);

  ASSERT_EQ(S_FALSE, enumNextDispID(0, did, &did));
}

//----------------------------------------------------------------------------
// partial data, after removing, iterate
TEST_F(PropertiesTest, IteratePartialPropertiesAfterRemove)
{
  setPropertyNames();
  setPropertyValues();

  EXPECT_HRESULT_SUCCEEDED(remove(testData[1].name));

  DISPID did = DISPID_STARTENUM;
  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[0].id, did);

  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[2].id, did);

  ASSERT_EQ(S_FALSE, enumNextDispID(0, did, &did));

  // set again - should be full set again
  setPropertyValues(1);

  did = DISPID_STARTENUM;
  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[0].id, did);

  ASSERT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[1].id, did);

  EXPECT_HRESULT_SUCCEEDED(enumNextDispID(0, did, &did));
  ASSERT_EQ(testData[2].id, did);

  ASSERT_EQ(S_FALSE, enumNextDispID(0, did, &did));
}
