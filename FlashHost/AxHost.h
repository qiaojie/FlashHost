#pragma once

class AxHost;

class AxHostWindow
{
public:
	virtual HWND GetAxHostWindow() const = 0;
	virtual void OnAxCreate(AxHost* host) = 0;
	virtual void OnAxInvalidate(const CRect& rect) = 0;
	virtual HRESULT QueryService(REFGUID guidService, REFIID riid, void** ppvObject) { return E_NOINTERFACE; }

protected:
	virtual ~AxHostWindow(){}
};

class AxHost :  public IDispatch,
				public IOleClientSite,
				public IOleContainer,
				public IOleControlSite,
				public IOleInPlaceSiteWindowless,
				public IAdviseSink,
				public IServiceProvider
{
public:
	AxHost(AxHostWindow* delegate);
	virtual ~AxHost();

	IUnknown* controlling_unknown();
	IUnknown* activex_control();

	bool CreateControl(const CLSID& clsid);
	bool SetRect(const CRect& rect);
	void Draw(HDC hdc, const CRect& rect);
	LRESULT OnWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT OnWindowlessMouseMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
	bool OnSetCursor(const CPoint& point);

	// IUnknown:
	STDMETHOD_(ULONG, AddRef)();
	STDMETHOD_(ULONG, Release)();
	STDMETHOD(QueryInterface)(REFIID iid, void** object);

	// IDispatch:
	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo);
	STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

	// IOleClientSite:
	STDMETHOD(SaveObject)();
	STDMETHOD(GetMoniker)(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk);
	STDMETHOD(GetContainer)(IOleContainer** ppContainer);
	STDMETHOD(ShowObject)();
	STDMETHOD(OnShowWindow)(BOOL fShow);
	STDMETHOD(RequestNewObjectLayout)();

	// IOleContainer:
	STDMETHOD(ParseDisplayName)(IBindCtx* pbc, LPOLESTR pszDisplayName, ULONG* pchEaten, IMoniker** ppmkOut);
	STDMETHOD(EnumObjects)(DWORD grfFlags, IEnumUnknown** ppenum);
	STDMETHOD(LockContainer)(BOOL fLock);

	// IOleControlSite:
	STDMETHOD(OnControlInfoChanged)();
	STDMETHOD(LockInPlaceActive)(BOOL fLock);
	STDMETHOD(GetExtendedControl)(IDispatch** ppDisp);
	STDMETHOD(TransformCoords)(POINTL* pPtlHimetric, POINTF* pPtfContainer, DWORD dwFlags);
	STDMETHOD(TranslateAccelerator)(MSG* pMsg, DWORD grfModifiers);
	STDMETHOD(OnFocus)(BOOL fGotFocus);
	STDMETHOD(ShowPropertyFrame)();

	// IOleWindow:
	STDMETHOD(GetWindow)(HWND* phwnd);
	STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode);

	// IOleInPlaceSite:
	STDMETHOD(CanInPlaceActivate)();
	STDMETHOD(OnInPlaceActivate)();
	STDMETHOD(OnUIActivate)();
	STDMETHOD(GetWindowContext)(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo);
	STDMETHOD(Scroll)(SIZE scrollExtant);
	STDMETHOD(OnUIDeactivate)(BOOL fUndoable);
	STDMETHOD(OnInPlaceDeactivate)();
	STDMETHOD(DiscardUndoState)();
	STDMETHOD(DeactivateAndUndo)();
	STDMETHOD(OnPosRectChange)(LPCRECT lprcPosRect);

	// IOleInPlaceSiteEx:
	STDMETHOD(OnInPlaceActivateEx)(BOOL* pfNoRedraw, DWORD dwFlags);
	STDMETHOD(OnInPlaceDeactivateEx)(BOOL fNoRedraw);
	STDMETHOD(RequestUIActivate)();

	// IOleInPlaceSiteWindowless:
	STDMETHOD(CanWindowlessActivate)();
	STDMETHOD(GetCapture)();
	STDMETHOD(SetCapture)(BOOL fCapture);
	STDMETHOD(GetFocus)();
	STDMETHOD(SetFocus)(BOOL fFocus);
	STDMETHOD(GetDC)(LPCRECT pRect, DWORD grfFlags, HDC* phDC);
	STDMETHOD(ReleaseDC)(HDC hDC);
	STDMETHOD(InvalidateRect)(LPCRECT pRect, BOOL fErase);
	STDMETHOD(InvalidateRgn)(HRGN hRGN, BOOL fErase);
	STDMETHOD(ScrollRect)(INT dx, INT dy, LPCRECT pRectScroll, LPCRECT pRectClip);
	STDMETHOD(AdjustRect)(LPRECT prc);
	STDMETHOD(OnDefWindowMessage)(UINT msg, WPARAM wParam, LPARAM lParam, LRESULT* plResult);

	// IAdviseSink:
	STDMETHOD_(void, OnDataChange)(FORMATETC* pFormatetc, STGMEDIUM* pStgmed);
	STDMETHOD_(void, OnViewChange)(DWORD dwAspect, LONG lindex);
	STDMETHOD_(void, OnRename)(IMoniker* pmk);
	STDMETHOD_(void, OnSave)();
	STDMETHOD_(void, OnClose)();

	// IServiceProvider:
	STDMETHOD(QueryService)(REFGUID guidService, REFIID riid, void** ppvObject);

private:
	bool ActivateAx();
	void ReleaseAll();

	AxHostWindow* _delegate;

	// IOleInPlaceSiteWindowless
	HDC _dc;
	bool _dcReleased;

	// pointers
	CComPtr<IUnknown>				_unknown;
	CComQIPtr<IOleObject>			_oleObject;
	CComQIPtr<IOleInPlaceFrame>		_inplaceFrame;
	CComQIPtr<IOleInPlaceUIWindow>	_inplaceUIWindow;
	CComQIPtr<IViewObjectEx>		_viewObject;
	CComQIPtr<IOleInPlaceObjectWindowless> _inplaceObjectWindowless;

	// state
	bool _inplaceActive;
	bool _uiActive;
	bool _windowless;
	bool _capture;
	bool _haveFocus;

	DWORD _oleObjectSink;
	DWORD _miscStatus;

	CRect _bounds;

	// Accelerator table
	HACCEL _accel;

	// Ambient property storage
	unsigned long _canWindowlessActivate : 1;
	unsigned long _userMode : 1;
};