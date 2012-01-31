// FlashHost.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "resource.h"
#include <process.h>
#include "AxHost.h"
#include "interface.h"
#include <gdiplus.h>

#pragma comment(lib, "gdiplus.lib")

using namespace Gdiplus;

#define WM_ACTION WM_USER + 2
#define WM_MOUSEMSG WM_USER + 3
#define WM_KEYMSG WM_USER + 4

class MainFrame : public CFrameWindowImpl<MainFrame>, public AxHostWindow, public CMessageLoop
{
public:
	DECLARE_FRAME_WND_CLASS(NULL, IDR_MAINFRAME)

	BEGIN_MSG_MAP(MainFrame)
	//ALT_MSG_MAP(0)
		MESSAGE_HANDLER(WM_CREATE,			OnWmCreate)
		MESSAGE_HANDLER(WM_ERASEBKGND,		OnEraseBkg)
		MESSAGE_HANDLER(WM_PAINT,           OnWmPaint)
		MESSAGE_HANDLER(WM_DESTROY,         OnClose)
		MESSAGE_HANDLER(WM_ACTION,			OnAction)
		MESSAGE_HANDLER(WM_MOUSEMSG,		OnMouseMsg)
		MESSAGE_HANDLER(WM_KEYMSG,			OnKeyMsg)
		MESSAGE_RANGE_HANDLER(WM_MOUSEFIRST, WM_MOUSELAST, OnWmMouseEvent)
		MESSAGE_RANGE_HANDLER(WM_KEYDOWN,   WM_CHAR, OnWmKeybEvent)
		CHAIN_MSG_MAP(CFrameWindowImpl<MainFrame>)
	END_MSG_MAP()


	// Message Handler
	LRESULT OnAction(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		IAction* action = (IAction*)wParam;
		action ->Run(_flash);
		return 0;
	}
	LRESULT OnMouseMsg(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		MSG* msg = (MSG*)wParam;
		_axHost->OnWindowlessMouseMessage(msg->message, msg->wParam, msg->lParam);
		return 0;
	}
	LRESULT OnKeyMsg(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		MSG* msg = (MSG*)wParam;
		_axHost->OnWindowMessage(msg->message, msg->wParam, msg->lParam);
		return 0;
	}
	
	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		delete _graph;
		delete _buffer;
		_flash->Stop();
		_flash.Release();
		delete _axHost;
		::PostQuitMessage(0);
		return 0;
	}
	LRESULT OnEraseBkg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		return 0;
	}
	LRESULT OnWmCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		_axHost->CreateControl(__uuidof(ShockwaveFlashObjects::ShockwaveFlash));
		_axHost->SetRect(CRect(0, 0, _bitmap.width, _bitmap.height));
		return 0;
	}

	LRESULT OnWmPaint(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
	{
		PAINTSTRUCT ps = { 0 };
		BeginPaint(&ps);
		{
			Graphics g(ps.hdc);
			g.DrawImage(_buffer, 0, 0);
		}
		EndPaint(&ps);
		return 0;
	}

	LRESULT OnWmMouseEvent(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
	{
		_axHost->OnWindowlessMouseMessage(uMsg, wParam, lParam);
		return 0;
	}

	LRESULT OnWmKeybEvent(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
	{
		_axHost->OnWindowMessage(uMsg, wParam, lParam);
		return 0;
	}

	// Overridden from AxHostWindow:
    virtual HWND GetAxHostWindow() const
	{
		return *this;
	}
    virtual void OnAxCreate(AxHost* host)
	{
		_flash = host->activex_control();
        HRESULT hr = _flash->put_WMode(L"Transparent");
	}
    virtual void OnAxInvalidate(const CRect& rect)
	{
		if(rect.Width() <= 0 || rect.Height() <= 0)
			return;

		_lock.Lock();

		_dirtyRects.push_back(rect);
		_bitmap.Clear(rect);
		HDC dc = _graph->GetHDC();
		_axHost->Draw(dc, CRect(0, 0, _bitmap.width, _bitmap.height));
		_graph->ReleaseHDC(dc);

		_lock.Unlock();
		InvalidateRect(&rect);
	}

	Graphics*	_graph;
	Bitmap*		_buffer;
	Bmp			_bitmap;
	AxHost*		_axHost;
	vector<CRect> _dirtyRects;
	CComQIPtr<ShockwaveFlashObjects::IShockwaveFlash> _flash;
	CComCriticalSection _lock;


	MainFrame()
	{
		_lock.Init();
		_axHost = new AxHost(this);
	}

	CComPtr<ShockwaveFlashObjects::IShockwaveFlash> GetFlash()
	{
		return _flash;
	}

	void SetClientSize(int w, int h)
	{
		_bitmap.width = w;
		_bitmap.height = h;
		_bitmap.pitch = w * 4;
		_bitmap.data = (BYTE*)malloc(_bitmap.pitch * _bitmap.height);
		_buffer = new Bitmap(w, h, _bitmap.pitch, PixelFormat32bppRGB, _bitmap.data);
		_graph = new Graphics(_buffer);
	}

	void Lock(Bmp* buffer, vector<CRect>& dirtyRects)
	{
		_lock.Lock();
		*buffer = _bitmap;
		dirtyRects = _dirtyRects;
	}

	void Unlock()
	{
		_dirtyRects.clear();
		_lock.Unlock();
	}
};


CAppModule _Module;

class FlashHost : public IFlashHost
{
	MainFrame _frameWnd;
	int _width, _height;
	bool _showWnd;
	HANDLE _creatFinish;
	HANDLE _thread;
	IStream* _flashObj;
	CComQIPtr<ShockwaveFlashObjects::IShockwaveFlash> _flash;

	static void ThreadFunc(void* arg)
	{
		((FlashHost*)arg)->ThreadFuncImpl();
	}

	void ThreadFuncImpl()
	{
		HRESULT hRes = ::CoInitialize(NULL);
		ATLASSERT(SUCCEEDED(hRes));
		::DefWindowProc(NULL, 0, 0, 0L);
		GdiplusStartupInput gdiplusStartupInput;
		ULONG_PTR gdiplusToken;
		GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
		hRes = _Module.Init(NULL, (HINSTANCE)&__ImageBase);
		ATLASSERT(SUCCEEDED(hRes));
		AtlAxWinInit();

		_frameWnd.SetClientSize(_width, _height);
		_frameWnd.Create();
		if(_showWnd)
			_frameWnd.ShowWindow(SW_SHOW);
	
		//hRes = ::CoMarshalInterThreadInterfaceInStream(__uuidof(ShockwaveFlashObjects::IShockwaveFlash), _frameWnd.GetFlash(), &_flashObj);
		//ATLASSERT(SUCCEEDED(hRes));
		::SetEvent(_creatFinish);

		_frameWnd.Run();

		_Module.Term();
		::CoUninitialize();
	}

	virtual void Destroy()
	{
		_frameWnd.SendMessage(WM_CLOSE, 0, 0);
		::WaitForSingleObject(_thread, -1);
		delete this;
	}

	virtual void Create(int width, int height, bool showWindow)
	{
		_creatFinish = ::CreateEvent(0,0,0,0);
		_width = width;
		_height = height;
		_showWnd = showWindow;

		_thread = (HANDLE)::_beginthread(ThreadFunc, 0, this);
		::WaitForSingleObject(_creatFinish, -1);

		//CComPtr<IUnknown> obj;
		//HRESULT hr = CoGetInterfaceAndReleaseStream(_flashObj, __uuidof(ShockwaveFlashObjects::IShockwaveFlash), (void**)&obj);
		//_flash = obj;
	}

	virtual void Lock(Bmp* buffer, vector<CRect>& dirtyRects)
	{
		_frameWnd.Lock(buffer, dirtyRects);
	}

	virtual void Unlock()
	{
		_frameWnd.Unlock();
	}

	virtual void DoAction(IAction* action)
	{
		_frameWnd.SendMessage(WM_ACTION, (WPARAM)action, 0);
	}

	virtual void SendMouseMessage(int message, WPARAM wParam, LPARAM lParam)
	{
		MSG msg;
		msg.message = message;
		msg.wParam = wParam;
		msg.lParam = lParam;
		_frameWnd.SendMessage(WM_MOUSEMSG, (WPARAM)&msg, 0);
	}

	virtual void SendKeyboardMessage(int message, WPARAM wParam, LPARAM lParam)
	{
		MSG msg;
		msg.message = message;
		msg.wParam = wParam;
		msg.lParam = lParam;
		_frameWnd.SendMessage(WM_KEYMSG, (WPARAM)&msg, 0);
	}

	virtual void GetWaveFormat(WAVEFORMATEX* format)
	{
	}

	virtual int GetSoundData(void* buf, int size)
	{
	}
};

IFlashHost* CreateFlashHost()
{
	return new FlashHost();
}
