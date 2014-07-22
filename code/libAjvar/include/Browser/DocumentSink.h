/**************************************************************************//**
 * @file
 * @brief     Declaration of DocumentSink
 * @author    Arne Seib <arne@salsitasoft.com>
 * @copyright 2014 Salsita Software (http://www.salsitasoft.com).
 *****************************************************************************/

#pragma once

namespace Ajvar {
namespace Browser {

//=============================================================================
/// @class   Ajvar::Browser::DocumentSink
/// @brief   Attaches to a HTML document and listens for readstate change events.
/// @details This class is a shortcut if you need notifications about document
/// ready state changes. Derive from this class and override the methods you need:
/// - OnDocumentLoaded
/// - OnDocumentInteractive
/// - OnDocumentComplete
template<class T>
  class DocumentSink :
    public IDispEventImpl<1, DocumentSink<T>, &DIID_HTMLDocumentEvents2, &LIBID_MSHTML, 4, 0>
{
public:
  /// @brief  Typedef for the baseclass.
  typedef IDispEventImpl<1, DocumentSink<T>, &DIID_HTMLDocumentEvents2, &LIBID_MSHTML, 4, 0> _EventImpl;

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

  BEGIN_SINK_MAP(DocumentSink<T>)
    SINK_ENTRY_EX(1, DIID_HTMLDocumentEvents2, DISPID_READYSTATECHANGE, _OnReadyStateChange)
  END_SINK_MAP()

  /// @brief  Event handler for `readyState` changes.
  STDMETHOD_(void, _OnReadyStateChange)(IHTMLEventObj* aEvent);

protected:
  /// @brief  Stop listening for `readyState` changes.
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
DocumentSink<T>::DocumentSink()
  : mReadyState(UNINITIALIZED)
{
}

// -------------------------------------------------------------------------
// destructor
template<class T>
DocumentSink<T>::~DocumentSink()
{
  UnadviseDocument();
}

// -------------------------------------------------------------------------
// _OnReadyStateChange
template<class T>
STDMETHODIMP_(void) DocumentSink<T>::_OnReadyStateChange(IHTMLEventObj* aEvent)
{
  FireCurrent(aEvent);
}

// -------------------------------------------------------------------------
// UnadviseDocument
template<class T>
void DocumentSink<T>::UnadviseDocument()
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
/// If `aRunCurrentState` is true, fire for the current state immediately.
/// If `FireCurrent()` returns true, the document is already in `complete`
/// state. In this case we are done, don't advise.
template<class T>
HRESULT DocumentSink<T>::AdviseDocument(IDispatch * aDispatch, bool aRunCurrentState)
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
/// @return : true if the document is complete or no document is available.
template<class T>
bool DocumentSink<T>::FireCurrent(IHTMLEventObj * aEvent = nullptr)
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

