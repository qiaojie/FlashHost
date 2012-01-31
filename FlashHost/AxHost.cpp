#include "stdafx.h"
#include "AxHost.h"


class AxFrameWindow : public CWindowImpl<AxFrameWindow>, public IOleInPlaceFrame
{
public:
	AxFrameWindow() : _ref(1) {}

	virtual ~AxFrameWindow()
	{
		_ax.Release();
		if(m_hWnd)
			DestroyWindow();
	}

	DECLARE_EMPTY_MSG_MAP()

	// IUnknown:
	STDMETHOD_(ULONG, AddRef)()
	{
		//assert_NE(_ref, -1);
		return InterlockedIncrement(&_ref);
	}

	STDMETHOD_(ULONG, Release)()
	{
		LONG ref = InterlockedDecrement(&_ref);
		if(ref == 0)
		{
			delete this;
		}
		return ref;
	}

	STDMETHOD(QueryInterface)(REFIID iid, void** object)
	{
		HRESULT hr = S_OK;
		*object = NULL;
		if(iid == IID_IUnknown || iid == IID_IOleWindow)
			*object = static_cast<IOleWindow*>(this);
		else if(iid == IID_IOleInPlaceUIWindow)
			*object = static_cast<IOleInPlaceUIWindow*>(this);
		else if(iid == IID_IOleInPlaceFrame)
			*object = static_cast<IOleInPlaceFrame*>(this);
		else
			hr = E_NOINTERFACE;

		if(SUCCEEDED(hr))
			static_cast<IUnknown*>(*object)->AddRef();

		return hr;
	}

	// IOleWindow
	STDMETHOD(GetWindow)(HWND* phwnd)
	{
		assert(phwnd);
		if(phwnd == NULL)
			return E_POINTER;

		if(m_hWnd == NULL)
			Create(NULL, NULL, _T("axhost frame window"),
			WS_OVERLAPPEDWINDOW, 0, (UINT)NULL);

		*phwnd = m_hWnd;
		return S_OK;
	}

	STDMETHOD(ContextSensitiveHelp)(BOOL /*fEnterMode*/)
	{
		return S_OK;
	}

	// IOleInPlaceUIWindow
	STDMETHOD(GetBorder)(LPRECT /*lprectBorder*/)
	{
		return S_OK;
	}

	STDMETHOD(RequestBorderSpace)(LPCBORDERWIDTHS /*pborderwidths*/)
	{
		return INPLACE_E_NOTOOLSPACE;
	}

	STDMETHOD(SetBorderSpace)(LPCBORDERWIDTHS /*pborderwidths*/)
	{
		return S_OK;
	}

	STDMETHOD(SetActiveObject)(IOleInPlaceActiveObject* pActiveObject,
		LPCOLESTR /*pszObjName*/)
	{
		_ax = pActiveObject;
		return S_OK;
	}

	// IOleInPlaceFrameWindow
	STDMETHOD(InsertMenus)(HMENU /*hmenuShared*/,
		LPOLEMENUGROUPWIDTHS /*lpMenuWidths*/)
	{
		return S_OK;
	}

	STDMETHOD(SetMenu)(HMENU /*hmenuShared*/,
		HOLEMENU /*holemenu*/,
		HWND /*hwndActiveObject*/)
	{
		return S_OK;
	}

	STDMETHOD(RemoveMenus)(HMENU /*hmenuShared*/)
	{
		return S_OK;
	}

	STDMETHOD(SetStatusText)(LPCOLESTR /*pszStatusText*/)
	{
		return S_OK;
	}

	STDMETHOD(EnableModeless)(BOOL /*fEnable*/)
	{
		return S_OK;
	}

	STDMETHOD(TranslateAccelerator)(LPMSG /*lpMsg*/, WORD /*wID*/)
	{
		return S_FALSE;
	}

private:
	LONG _ref;
	CComPtr<IOleInPlaceActiveObject> _ax;
};


class AxUIWindow : public CWindowImpl<AxUIWindow>, public IOleInPlaceUIWindow
{
public:
	AxUIWindow() : _ref(1) {}
	virtual ~AxUIWindow()
	{
		_ax.Release();
		if(m_hWnd)
			DestroyWindow();
	}

	DECLARE_EMPTY_MSG_MAP()

	// IUnknown:
	STDMETHOD_(ULONG, AddRef)()
	{
		//assert_NE(_ref, -1);
		return InterlockedIncrement(&_ref);
	}

	STDMETHOD_(ULONG, Release)()
	{
		LONG ref = InterlockedDecrement(&_ref);
		if(ref == 0)
		{
			delete this;
		}
		return ref;
	}

	STDMETHOD(QueryInterface)(REFIID iid, void** object)
	{
		HRESULT hr = S_OK;
		*object = NULL;
		if(iid == IID_IUnknown || iid == IID_IOleWindow)
			*object = static_cast<IOleWindow*>(this);
		else if(iid == IID_IOleInPlaceUIWindow)
			*object = static_cast<IOleInPlaceUIWindow*>(this);
		else
			hr = E_NOINTERFACE;

		if(SUCCEEDED(hr))
			static_cast<IUnknown*>(*object)->AddRef();
		return hr;
	}

	// IOleWindow
	STDMETHOD(GetWindow)(HWND* phwnd)
	{
		assert(phwnd);
		if(phwnd == NULL)
			return E_POINTER;

		if(m_hWnd == NULL)
			Create(NULL, NULL, _T("axhost ui window"), WS_OVERLAPPEDWINDOW, 0, (UINT)NULL);

		*phwnd = m_hWnd;
		return S_OK;
	}

	STDMETHOD(ContextSensitiveHelp)(BOOL /*fEnterMode*/)
	{
		return S_OK;
	}

	// IOleInPlaceUIWindow
	STDMETHOD(GetBorder)(LPRECT /*lprectBorder*/)
	{
		return S_OK;
	}

	STDMETHOD(RequestBorderSpace)(LPCBORDERWIDTHS /*pborderwidths*/)
	{
		return INPLACE_E_NOTOOLSPACE;
	}

	STDMETHOD(SetBorderSpace)(LPCBORDERWIDTHS /*pborderwidths*/)
	{
		return S_OK;
	}

	STDMETHOD(SetActiveObject)(IOleInPlaceActiveObject* pActiveObject, LPCOLESTR /*pszObjName*/)
	{
		_ax = pActiveObject;
		return S_OK;
	}

private:
	LONG _ref;
	CComPtr<IOleInPlaceActiveObject> _ax;
};




AxHost::AxHost(AxHostWindow* del)
	: _delegate(del),
	_dc(NULL),
	_dcReleased(true),
	_inplaceActive(false),
	_uiActive(false),
	_windowless(false),
	_capture(false),
	_haveFocus(false),
	_oleObjectSink(0),
	_miscStatus(0),
	_accel(NULL),
	_canWindowlessActivate(true),
	_userMode(true)
{
	assert(_delegate);
}

AxHost::~AxHost()
{
	ReleaseAll();
}

IUnknown* AxHost::controlling_unknown()
{
	return static_cast<IDispatch*>(this);
}

IUnknown* AxHost::activex_control()
{
	return _unknown;
}

bool AxHost::CreateControl(const CLSID& clsid)
{
	ReleaseAll();

	HRESULT hr = _unknown.CoCreateInstance(clsid);
	if(FAILED(hr))
		return false;

	bool success = ActivateAx();
	if(!success)
		ReleaseAll();

	return success;
}

bool AxHost::SetRect(const CRect& rect)
{
	if(!::EqualRect(&_bounds, &rect))
	{
		_bounds = rect;
		SIZEL pxsize = { _bounds.right - _bounds.left, _bounds.bottom - _bounds.top };
		SIZEL hmsize = { 0 };
		AtlPixelToHiMetric(&pxsize, &hmsize);

		if(_oleObject)
		{
			_oleObject->SetExtent(DVASPECT_CONTENT, &hmsize);
			_oleObject->GetExtent(DVASPECT_CONTENT, &hmsize);
			AtlHiMetricToPixel(&hmsize, &pxsize);
			_bounds.right = _bounds.left + pxsize.cx;
			_bounds.bottom = _bounds.top + pxsize.cy;
		}
		if(_inplaceObjectWindowless)
			_inplaceObjectWindowless->SetObjectRects(&_bounds, &_bounds);

		if(_windowless)
			_delegate->OnAxInvalidate(_bounds);
	}

	return true;
}

void AxHost::Draw(HDC hdc, const CRect& rect)
{
	if(_viewObject && _windowless)
		OleDraw(_viewObject, DVASPECT_CONTENT, hdc, &rect);
}

LRESULT AxHost::OnWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lRes = 0;
	if(_inplaceActive && _windowless && _inplaceObjectWindowless)
		_inplaceObjectWindowless->OnWindowMessage(uMsg, wParam, lParam, &lRes);

	return lRes;
}

LRESULT AxHost::OnWindowlessMouseMessage(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lRes = 0;
	if(_inplaceActive && _windowless && _inplaceObjectWindowless)
		_inplaceObjectWindowless->OnWindowMessage(uMsg, wParam, lParam, &lRes);

	return lRes;
}

bool AxHost::OnSetCursor(const CPoint& point)
{
	if(!_windowless)
		return false;

	DWORD dwHitResult = _capture ? HITRESULT_HIT : HITRESULT_OUTSIDE;
	if(dwHitResult==HITRESULT_OUTSIDE && _viewObject)
		_viewObject->QueryHitPoint(DVASPECT_CONTENT, &_bounds, point, 0, &dwHitResult);

	if(dwHitResult == HITRESULT_HIT)
		return true;

	return false;
}

// IUnknown:
STDMETHODIMP_(ULONG) AxHost::AddRef()
{
	return 1;
}

STDMETHODIMP_(ULONG) AxHost::Release()
{
	return 1;
}

STDMETHODIMP AxHost::QueryInterface(REFIID iid, void** object)
{
	HRESULT hr = S_OK;
	*object = NULL;
	if(iid==IID_IUnknown || iid==IID_IDispatch)
		*object = controlling_unknown();
	else if(iid == IID_IOleClientSite)
		*object = static_cast<IOleClientSite*>(this);
	else if(iid == IID_IOleContainer)
		*object = static_cast<IOleContainer*>(this);
	else if(iid == IID_IOleControlSite)
		*object = static_cast<IOleControlSite*>(this);
	else if(iid == IID_IOleWindow)
		*object = static_cast<IOleWindow*>(this);
	else if(iid == IID_IOleInPlaceSite)
		*object = static_cast<IOleInPlaceSite*>(this);
	else if(iid == IID_IOleInPlaceSiteEx)
		*object = static_cast<IOleInPlaceSiteEx*>(this);
	else if(iid == IID_IOleInPlaceSiteWindowless)
		*object = static_cast<IOleInPlaceSiteWindowless*>(this);
	else if(iid == IID_IAdviseSink)
		*object = static_cast<IAdviseSink*>(this);
	else if(iid == IID_IServiceProvider)
		*object = static_cast<IServiceProvider*>(this);
	else
		hr = E_NOINTERFACE;

	if(SUCCEEDED(hr))
		static_cast<IUnknown*>(*object)->AddRef();

	return hr;
}

// IDispatch:
STDMETHODIMP AxHost::GetTypeInfoCount(UINT* pctinfo)
{
	if(pctinfo == NULL) 
		return E_POINTER; 

	*pctinfo = 1;
	return S_OK;
}

STDMETHODIMP AxHost::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags,
							DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	HRESULT hr;
	VARIANT varResult;

	if(riid != IID_NULL)
		return E_INVALIDARG;

	if(NULL == pVarResult)
		pVarResult = &varResult;

	VariantInit(pVarResult);

	//set default as boolean
	V_VT(pVarResult)=VT_BOOL;

	// implementation is read-only
	if(!(DISPATCH_PROPERTYGET & wFlags))
		return DISP_E_MEMBERNOTFOUND;

	hr = S_OK;
	switch(dispIdMember)
	{
	case DISPID_AMBIENT_USERMODE:
		V_BOOL(pVarResult) = _userMode ? VARIANT_TRUE : VARIANT_FALSE;
		break;

	case DISPID_AMBIENT_UIDEAD:
		V_BOOL(pVarResult) = VARIANT_FALSE;
		break;

	case DISPID_AMBIENT_SUPPORTSMNEMONICS:
		V_BOOL(pVarResult) = VARIANT_FALSE;
		break;

	case DISPID_AMBIENT_SHOWGRABHANDLES:
		V_BOOL(pVarResult) = (!_userMode) ? VARIANT_TRUE : VARIANT_FALSE;
		break;

	case DISPID_AMBIENT_SHOWHATCHING:
		V_BOOL(pVarResult) = (!_userMode) ? VARIANT_TRUE : VARIANT_FALSE;
		break;

	default:
		hr = DISP_E_MEMBERNOTFOUND;
		break;
	}
	return hr;
}

// IOleClientSite:
STDMETHODIMP AxHost::SaveObject()
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::GetContainer(IOleContainer** ppContainer)
{
	assert(ppContainer);

	HRESULT hr = E_POINTER;
	if(ppContainer)
	{
		hr = E_NOTIMPL;
		(*ppContainer) = NULL;

		if(FAILED(hr))
			hr = QueryInterface(__uuidof(IOleContainer), (void**)ppContainer);
	}

	return hr;
}

STDMETHODIMP AxHost::ShowObject()
{
	_delegate->OnAxInvalidate(CRect(_bounds));
	return S_OK;
}

STDMETHODIMP AxHost::OnShowWindow(BOOL fShow)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::RequestNewObjectLayout()
{
	return E_NOTIMPL;
}

// IOleContainer:
STDMETHODIMP AxHost::ParseDisplayName(IBindCtx* pbc, LPOLESTR pszDisplayName, ULONG* pchEaten, IMoniker** ppmkOut)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::EnumObjects(DWORD grfFlags, IEnumUnknown** ppenum)
{
	if(ppenum == NULL)
	{
		return E_POINTER;
	}
	*ppenum = NULL;

	return E_NOTIMPL;
}

STDMETHODIMP AxHost::LockContainer(BOOL fLock)
{
	return S_OK;
}

// IOleControlSite:
STDMETHODIMP AxHost::OnControlInfoChanged()
{
	return S_OK;
}

STDMETHODIMP AxHost::LockInPlaceActive(BOOL fLock)
{
	return S_OK;
}

STDMETHODIMP AxHost::GetExtendedControl(IDispatch** ppDisp)
{
	if(ppDisp == NULL)
		return E_POINTER;

	return _oleObject.QueryInterface(ppDisp);
}

STDMETHODIMP AxHost::TransformCoords(POINTL* pPtlHimetric, POINTF* pPtfContainer, DWORD dwFlags)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::TranslateAccelerator(MSG* pMsg, DWORD grfModifiers)
{
	return S_FALSE;
}

STDMETHODIMP AxHost::OnFocus(BOOL fGotFocus)
{
	_haveFocus = fGotFocus;
	return S_OK;
}

STDMETHODIMP AxHost::ShowPropertyFrame()
{
	return E_NOTIMPL;
}

// IOleWindow:
STDMETHODIMP AxHost::GetWindow(HWND* phwnd)
{
	assert(phwnd);
	*phwnd = _delegate->GetAxHostWindow();
	return S_OK;
}

STDMETHODIMP AxHost::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

// IOleInPlaceSite:
STDMETHODIMP AxHost::CanInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP AxHost::OnInPlaceActivate()
{
	assert(!_inplaceActive);
	assert(!_inplaceObjectWindowless);

	_inplaceActive = true;
	OleLockRunning(_oleObject, TRUE, FALSE);
	_windowless = false;
	_oleObject->QueryInterface(__uuidof(IOleInPlaceObject),
		(void**)&_inplaceObjectWindowless);
	return S_OK;
}

STDMETHODIMP AxHost::OnUIActivate()
{
	_uiActive = true;
	return S_OK;
}

STDMETHODIMP AxHost::GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc,
									  LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	if(ppFrame != NULL)
		*ppFrame = NULL;
	if(ppDoc != NULL)
		*ppDoc = NULL;
	if(ppFrame==NULL || ppDoc==NULL || lprcPosRect==NULL || lprcClipRect==NULL)
		return E_POINTER;

	if(!_inplaceFrame)
	{
		_inplaceFrame.Attach(new AxFrameWindow());
		assert(_inplaceFrame);
	}
	if(!_inplaceUIWindow)
	{
		_inplaceUIWindow.Attach(new AxUIWindow());
		assert(_inplaceUIWindow);
	}

	HRESULT hr = _inplaceFrame.QueryInterface(ppFrame);
	if(FAILED(hr))
		return hr;

	hr = _inplaceUIWindow.QueryInterface(ppDoc);
	if(FAILED(hr))
		return hr;

	*lprcPosRect = _bounds;
	*lprcClipRect = _bounds;

	if(_accel == NULL)
	{
		ACCEL ac = { 0, 0, 0 };
		_accel = CreateAcceleratorTable(&ac, 1);
	}
	lpFrameInfo->cb = sizeof(OLEINPLACEFRAMEINFO);
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = _delegate->GetAxHostWindow();
	lpFrameInfo->haccel = _accel;
	lpFrameInfo->cAccelEntries = (_accel != NULL) ? 1 : 0;

	return hr;
}

STDMETHODIMP AxHost::Scroll(SIZE scrollExtant)
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::OnUIDeactivate(BOOL fUndoable)
{
	_uiActive = false;
	return S_OK;
}

STDMETHODIMP AxHost::OnInPlaceDeactivate()
{
	return OnInPlaceDeactivateEx(TRUE);
}

STDMETHODIMP AxHost::DiscardUndoState()
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::DeactivateAndUndo()
{
	return E_NOTIMPL;
}

STDMETHODIMP AxHost::OnPosRectChange(LPCRECT lprcPosRect)
{
	if(lprcPosRect == 0)
		return E_POINTER;

	return S_OK;
}

// IOleInPlaceSiteEx:
STDMETHODIMP AxHost::OnInPlaceActivateEx(BOOL* /*pfNoRedraw*/, DWORD dwFlags)
{
	assert(!_inplaceActive);
	assert(!_inplaceObjectWindowless);

	_inplaceActive = true;
	OleLockRunning(_oleObject, TRUE, FALSE);

	HRESULT hr = E_FAIL;
	if(dwFlags & ACTIVATE_WINDOWLESS)
	{
		_windowless = true;
		_inplaceObjectWindowless = _oleObject;
	}
	if(!_inplaceObjectWindowless)
	{
		_windowless = false;
		hr = _oleObject->QueryInterface(__uuidof(IOleInPlaceObject), (void**)&_inplaceObjectWindowless);
	}
	if(_inplaceObjectWindowless)
		_inplaceObjectWindowless->SetObjectRects(&_bounds, &_bounds);

	return hr;
}

STDMETHODIMP AxHost::OnInPlaceDeactivateEx(BOOL fNoRedraw)
{
	_inplaceActive = false;
	_inplaceObjectWindowless.Release();
	return S_OK;
}

STDMETHODIMP AxHost::RequestUIActivate()
{
	return S_OK;
}

// IOleInPlaceSiteWindowless:
STDMETHODIMP AxHost::CanWindowlessActivate()
{
	return _canWindowlessActivate ? S_OK : S_FALSE;
}

STDMETHODIMP AxHost::GetCapture()
{
	return _capture ? S_OK : S_FALSE;
}

STDMETHODIMP AxHost::SetCapture(BOOL fCapture)
{
	if(fCapture)
		_capture = true;
	else
		_capture = false;

	return S_OK;
}

STDMETHODIMP AxHost::GetFocus()
{
	return _haveFocus ? S_OK : S_FALSE;
}

STDMETHODIMP AxHost::SetFocus(BOOL fFocus)
{
	_haveFocus = fFocus;
	return S_OK;
}

STDMETHODIMP AxHost::GetDC(LPCRECT pRect, DWORD grfFlags, HDC* phDC)
{
	if(phDC == NULL)
		return E_POINTER;
	if(!_dcReleased)
		return E_FAIL;

	*phDC = ::GetDC(_delegate->GetAxHostWindow());
	if(*phDC == NULL)
		return E_FAIL;

	_dcReleased = false;

	if(grfFlags & OLEDC_NODRAW)
		return S_OK;

	if(grfFlags & OLEDC_OFFSCREEN)
	{
		HDC hDCOffscreen = CreateCompatibleDC(*phDC);
		if(hDCOffscreen != NULL)
		{
			HBITMAP hBitmap = CreateCompatibleBitmap(*phDC, _bounds.Width(), _bounds.Height());

			if(hBitmap == NULL)
				DeleteDC(hDCOffscreen);
			else
			{
				HGDIOBJ hOldBitmap = SelectObject(hDCOffscreen, hBitmap);
				if(hOldBitmap == NULL)
				{
					DeleteObject(hBitmap);
					DeleteDC(hDCOffscreen);
				}
				else
				{
					DeleteObject(hOldBitmap);
					_dc = *phDC;
					*phDC = hDCOffscreen;
				}
			}
		}
	}

	return S_OK;
}

STDMETHODIMP AxHost::ReleaseDC(HDC hDC)
{
	_dcReleased = true;
	if(_dc != NULL)
	{
		// Offscreen DC has to be copied to screen DC before releasing the screen dc;
		BitBlt(_dc, _bounds.left, _bounds.top,
			_bounds.right-_bounds.left,
			_bounds.bottom-_bounds.top,
			hDC, 0, 0, SRCCOPY);
		DeleteDC(hDC);
		hDC = _dc;
	}

	::ReleaseDC(_delegate->GetAxHostWindow(), hDC);
	return S_OK;
}

STDMETHODIMP AxHost::InvalidateRect(LPCRECT pRect, BOOL fErase)
{
	RECT rect = { 0 };
	if(pRect == NULL)
	{
		rect = _bounds;
	}
	else
	{
		IntersectRect(&rect, pRect, &_bounds);
	}

	if(!IsRectEmpty(&rect))
	{
		_delegate->OnAxInvalidate(CRect(rect));
	}

	return S_OK;
}

STDMETHODIMP AxHost::InvalidateRgn(HRGN hRGN, BOOL fErase)
{
	if(hRGN == NULL)
	{
		return InvalidateRect(NULL, fErase);
	}

	RECT rc_rgn = { 0 };
	GetRgnBox(hRGN, &rc_rgn);
	return InvalidateRect(&rc_rgn, fErase);
}

STDMETHODIMP AxHost::ScrollRect(INT dx, INT dy,
								LPCRECT pRectScroll,
								LPCRECT pRectClip)
{
	return S_OK;
}

STDMETHODIMP AxHost::AdjustRect(LPRECT prc)
{
	return S_OK;
}

STDMETHODIMP AxHost::OnDefWindowMessage(UINT msg,
										WPARAM wParam,
										LPARAM lParam,
										LRESULT* plResult)
{
	*plResult = DefWindowProc(_delegate->GetAxHostWindow(), msg, wParam, lParam);
	return S_OK;
}

// IAdviseSink:
STDMETHODIMP_(void) AxHost::OnDataChange(FORMATETC* pFormatetc, STGMEDIUM* pStgmed) {}

STDMETHODIMP_(void) AxHost::OnViewChange(DWORD dwAspect, LONG lindex) {}

STDMETHODIMP_(void) AxHost::OnRename(IMoniker* pmk) {}

STDMETHODIMP_(void) AxHost::OnSave() {}

STDMETHODIMP_(void) AxHost::OnClose() {}

// IServiceProvider:
STDMETHODIMP AxHost::QueryService(REFGUID guidService,
								  REFIID riid, void** ppvObject)
{
	assert(ppvObject);
	if(ppvObject == NULL)
	{
		return E_POINTER;
	}
	*ppvObject = NULL;

	HRESULT hr = E_NOINTERFACE;
	// Try for service on this object

	// No services currently
	if(FAILED(hr))
	{
		hr = _delegate->QueryService(guidService, riid, ppvObject);
	}

	return hr;
}

// private:
bool AxHost::ActivateAx()
{
	_oleObject = _unknown;
	if(!_oleObject)
		return false;

	_oleObject->GetMiscStatus(DVASPECT_CONTENT, &_miscStatus);
	if(_miscStatus & OLEMISC_SETCLIENTSITEFIRST)
		_oleObject->SetClientSite(this);

	CComQIPtr<IPersistStreamInit> stream_init;
	stream_init = _oleObject;
	HRESULT hr = stream_init->InitNew();
	if(FAILED(hr))
	{
		// Clean up and return
		if(_miscStatus & OLEMISC_SETCLIENTSITEFIRST)
			_oleObject->SetClientSite(0);

		_miscStatus = 0;
		_oleObject.Release();
		_unknown.Release();

		return false;
	}

	if(0 == (_miscStatus & OLEMISC_SETCLIENTSITEFIRST))
		_oleObject->SetClientSite(this);

	_delegate->OnAxCreate(this);

	_viewObject = _oleObject;
	if(!_viewObject)
	{
		hr = _oleObject->QueryInterface(__uuidof(IViewObject2), (void**)&_viewObject);
		if(FAILED(hr))
			hr = _oleObject->QueryInterface(__uuidof(IViewObject), (void**)&_viewObject);
	}


	CComQIPtr<IAdviseSink> advise_sink = controlling_unknown();
	_oleObject->Advise(advise_sink, &_oleObjectSink);
	if(_viewObject)
		_viewObject->SetAdvise(DVASPECT_CONTENT, 0, this);

	_oleObject->SetHostNames(L"axhost", NULL);
	if((_miscStatus & OLEMISC_INVISIBLEATRUNTIME) == 0)
	{
        hr = _oleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, this, 0, _delegate->GetAxHostWindow(), &_bounds);
	}

	CComQIPtr<IObjectWithSite> site = _unknown;
	if(site)
		site->SetSite(controlling_unknown());

	return true;
}

void AxHost::ReleaseAll()
{
	if(_viewObject)
	{
		_viewObject->SetAdvise(DVASPECT_CONTENT, 0, NULL);
		_viewObject.Release();
	}

	if(_oleObject)
	{
		_oleObject->Unadvise(_oleObjectSink);
		_oleObject->Close(OLECLOSE_NOSAVE);
		_oleObject->SetClientSite(NULL);
		_oleObject.Release();
	}

	if(_unknown)
	{
		CComQIPtr<IObjectWithSite> site = _unknown;
		if(site)
			site->SetSite(NULL);
		_unknown.Release();
	}

	_inplaceObjectWindowless.Release();
	_inplaceFrame.Release();
	_inplaceUIWindow.Release();

	_inplaceActive = false;
	_uiActive = false;
	_windowless = false;
	_capture = false;
	_haveFocus = false;

	if(_accel != NULL)
	{
		DestroyAcceleratorTable(_accel);
		_accel = NULL;
	}
}
