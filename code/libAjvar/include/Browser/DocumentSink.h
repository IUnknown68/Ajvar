/**************************************************************************//**
 * @file
 * @brief     Declaration of DocumentSink
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/

#pragma once

namespace Ajvar {

/// @brief This and that for the webbrowser control.
namespace Browser {

//=============================================================================
/// @class   Ajvar::Browser::DocumentSink
/// @brief   Attaches to a HTML document and listens for readstate change events.
/// @details This class is a shortcut if you need notifications about document
/// ready state changes. Derive from this class and override the methods you need:
/// - `OnDocumentLoading(IHTMLEventObj * aEvent)`
/// - `OnDocumentLoaded(IHTMLEventObj * aEvent)`
/// - `OnDocumentInteractive(IHTMLEventObj * aEvent)`
/// - `OnDocumentComplete(IHTMLEventObj * aEvent)`
template<UINT nID, class T>
  class DocumentSink :
    public IDispEventImpl<nID, DocumentSink<nID, T>, &DIID_HTMLDocumentEvents2, &LIBID_MSHTML, 4, 0>
{
public:
  /// @brief  Typedef for the `IDispEventImpl`.
  typedef IDispEventImpl<nID, DocumentSink<nID, T>, &DIID_HTMLDocumentEvents2, &LIBID_MSHTML, 4, 0> _EventImpl;

  /// @brief  Numeric `readyState` values.
  enum ReadyState {
    UNINITIALIZED = 0,
    LOADING,
    LOADED,
    INTERACTIVE,
    COMPLETE
  };

public:
  DocumentSink();
  virtual ~DocumentSink();

  BEGIN_SINK_MAP(DocumentSink<nID, T>)
    SINK_ENTRY_EX(1, DIID_HTMLDocumentEvents2, DISPID_READYSTATECHANGE, _OnReadyStateChange)
  END_SINK_MAP()

  /// @brief  Event handler for `readyState` changes.
  STDMETHOD_(void, _OnReadyStateChange)(IHTMLEventObj* aEvent);

protected:
  /// @brief  Stop listening for `readyState` changes and detach.
  void UnadviseDocument();

  /// @brief  Start listening for `readyState` changes.
  HRESULT AdviseDocument(IDispatch * aDispatch, bool aRunCurrentState);

  /// @brief  Fires for current readystate.
  /// @return true if the document is complete or no document is available.
  bool FireCurrent(IHTMLEventObj * aEvent = nullptr);

protected:
  /// @brief  The current HTML document.
  CComPtr<IHTMLDocument2> mDocument;

  int  mReadyState;
  
};

//=============================================================================
//=============================================================================
// Implementation

// -------------------------------------------------------------------------
// constructor
template<class T>
DocumentSink<nID, T>::DocumentSink()
  : mReadyState(UNINITIALIZED)
{
}

// -------------------------------------------------------------------------
// destructor
template<UINT nID, class T>
DocumentSink<nID, T>::~DocumentSink()
{
  UnadviseDocument();
}

// -------------------------------------------------------------------------
// _OnReadyStateChange
template<UINT nID, class T>
STDMETHODIMP_(void) DocumentSink<nID, T>::_OnReadyStateChange(IHTMLEventObj* aEvent)
{
  FireCurrent(aEvent);
}

// -------------------------------------------------------------------------
// UnadviseDocument
template<UINT nID, class T>
void DocumentSink<nID, T>::UnadviseDocument()
{
  if (mDocument) {
    _EventImpl::DispEventUnadvise(mDocument);
    mDocument = nullptr;
  }

}

// -------------------------------------------------------------------------
// AdviseDocument
/// @param    aDispatch         Dispatch pointer containing the HTML document  
///                             **or** a IWebBrowser2 where to get the  
///                             document from.
/// @param    aRunCurrentState  If true, invoke event for current state.
/// @details  Attaches to document events via `DispEventAdvise`. If there is
/// already a document, unadvise first.  
///
/// If `aRunCurrentState` is true, fire for the current state immediately.
///
/// If `FireCurrent()` returns true, the document is already in `complete`
/// state. In this case we are done, don't advise.
template<UINT nID, class T>
HRESULT DocumentSink<nID, T>::AdviseDocument(IDispatch * aDispatch, bool aRunCurrentState)
{
  if (mDocument) {
    UnadviseDocument();
  }
  CComQIPtr<IHTMLDocument2> document(aDispatch);
  if (nullptr == document) {
    CComQIPtr<IWebBrowser2> browser(aDispatch);
    if (nullptr == browser) {
      return E_NOINTERFACE;
    }
    CComPtr<IDispatch> documentDispatch;
    browser->get_Document(&documentDispatch);
    document = documentDispatch;
    if (nullptr == document) {
      return E_NOINTERFACE;
    }
  }

  if (aRunCurrentState) {
    // immediately run callback for current state
    mDocument = document;
    bool complete = FireCurrent(nullptr);
    mDocument = nullptr;
    if (complete) {
      // complete? so we are done.
      return S_OK;
    }
  }

  // advise for later
  HRESULT hr = _EventImpl::DispEventAdvise(document);
  if (SUCCEEDED(hr)) {
    mDocument = document;
  }
  return hr;
}

//----------------------------------------------------------------------------
// FireCurrent
/// @return true if the document is complete or no document is available.
template<UINT nID, class T>
bool DocumentSink<nID, T>::FireCurrent(IHTMLEventObj * aEvent = nullptr)
{
  if (!mDocument) {
    return true;
  }

  int newState = UNINITIALIZED;

  CComBSTR readyState;
  mDocument->get_readyState(&readyState);
  if (readyState == L"loading") {
    newState = LOADING;
  }
  else if (readyState == L"loaded") {
    newState = LOADED;
  }
  else if (readyState == L"interactive") {
    newState = INTERACTIVE;
  }
  else if (readyState == L"complete") {
    newState = COMPLETE;
  }

  bool complete = false;
  for (mReadyState; mReadyState <= newState; mReadyState++) {
    switch(mReadyState) {
      case LOADING:
        __if_exists(T::OnDocumentLoading) {
          static_cast<T*>(this)->OnDocumentLoading(aEvent);
        }
        break;
      case LOADED:
        __if_exists(T::OnDocumentLoaded) {
          static_cast<T*>(this)->OnDocumentLoaded(aEvent);
        }
        break;
      case INTERACTIVE:
        __if_exists(T::OnDocumentInteractive) {
          static_cast<T*>(this)->OnDocumentInteractive(aEvent);
        }
        break;
      case COMPLETE:
        __if_exists(T::OnDocumentComplete) {
          static_cast<T*>(this)->OnDocumentComplete(aEvent);
        }
        complete = true;
        break;
    }
  }
  return complete;
}

} // namespace Browser
} // namespace Ajvar

