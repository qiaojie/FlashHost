#include "stdafx.h"
#include "interface.h"



int _tmain(int argc, _TCHAR* argv[])
{
	::CoInitialize(NULL);

	HRESULT hr;
	IFlashHost* host = ::CreateFlashHost();
	host->Create(1024, 768, true);
	//hr = host->GetFlash()->put_Movie(L"D:\\projects\\FlashHost\\test.swf");
	//hr = host->GetFlash()->Play();
	struct PlayAction : IAction
	{
		void Run(ShockwaveFlashObjects::IShockwaveFlash* flash)
		{
			flash->put_Movie(L"D:\\projects\\FlashHost\\test.swf");
			flash->Play();
		}
	};

	PlayAction play;
	host->DoAction(&play);

	while(true)
	{
		Bmp buffer;
		vector<CRect> dirtyRects;
		host->Lock(&buffer, dirtyRects);
		if(dirtyRects.size() > 0)
		{
			printf("update:\n");
			for(int i = 0; i < dirtyRects.size(); ++i)
				printf("  (%d,%d,%d,%d)", dirtyRects[i].left, dirtyRects[i].top, dirtyRects[i].Width(), dirtyRects[i].Height());
			printf("\n");
		}
		host->Unlock();
		::Sleep(40);
	}

	host->Destroy();	
	::CoUninitialize();
}