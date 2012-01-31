#include "stdafx.h"
#include "interface.h"



int _tmain(int argc, _TCHAR* argv[])
{
	::CoInitialize(NULL);

	HRESULT hr;
	IFlashHost* host = ::CreateFlashHost();
	host->Create(1024, 768, true);

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
	WAVEFORMATEX format;
	::Sleep(100);
	host->GetWaveFormat(&format);
	FILE* waveFile = fopen("test.wav", "wb");
	int size = 16;
	fwrite("RIFF", 1, 4, waveFile);
	fwrite(&size, 1, 4, waveFile);
	fwrite("WAVE", 1, 4, waveFile);
	fwrite("fmt ", 1, 4, waveFile);
	fwrite(&size, 1, 4, waveFile);
	fwrite(&format, 1, 16, waveFile);
	fwrite("data", 1, 4, waveFile);
	fwrite(&size, 1, 4, waveFile);

	int soundLen = 0;

	for(int i = 0; i < 200; ++i)
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
		BYTE sndBuf[65536];
		int len = host->GetSoundData(sndBuf, 65536);
		fwrite(sndBuf, 1, len, waveFile);
		soundLen += len;
		host->Unlock();
		::Sleep(40);
	}

	size = soundLen + 36;
	fseek(waveFile, 4, SEEK_SET);
	fwrite(&size, 1, 4, waveFile);
	size = soundLen;
	fseek(waveFile, 40, SEEK_SET);
	fwrite(&size, 1, 4, waveFile);
	fclose(waveFile);

	host->Destroy();	
	::CoUninitialize();
}