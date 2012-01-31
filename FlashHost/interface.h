#pragma once
#import "libid:D27CDB6B-AE6D-11CF-96B8-444553540000"
#include <vector>
using namespace std;

struct Bmp
{
	int width, height;
	int pitch;
	BYTE* data;

	void Clear(const CRect& rect)
	{
		for(int i = rect.top; i < rect.bottom; ++i)
			ZeroMemory(data + i * pitch + rect.left * 4, rect.Width() * 4);
	}

	DWORD* operator[](int y)
	{
		return (DWORD*)(data + y * pitch);
	}
};

struct IAction
{
	virtual void Run(ShockwaveFlashObjects::IShockwaveFlash* flash) = 0;
};

struct IFlashHost
{
	virtual void Destroy() = 0;

	virtual void Create(int width, int height, bool showWindow) = 0;

	virtual void Lock(Bmp* buffer, vector<CRect>& dirtyRects) = 0;

	virtual void Unlock() = 0;

	virtual void DoAction(IAction* action) = 0;

	virtual void SendMouseMessage(int message, WPARAM wParam, LPARAM lParam) = 0;

	virtual void SendKeyboardMessage(int message, WPARAM wParam, LPARAM lParam) = 0;

	virtual void GetWaveFormat(WAVEFORMATEX* format) = 0;

	virtual int GetSoundData(void* buf, int size) = 0;
};

IFlashHost* CreateFlashHost();