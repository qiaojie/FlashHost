// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include <atlbase.h>
#include "../detours/detours.h"
#include <mmsystem.h>
#include <string>
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "../detours/detours.lib")

template<int Size = 256 * 1024>
struct RingBuffer
{
	char buf[Size];
	volatile int start;
	volatile int end;

	RingBuffer()
	{
		start = 0;
		end = 0;
	}

	int Write(char* data, int size)
	{
		int remain;
		if(end < start)
			remain = Size + end - start - 1;
		else
			remain = Size - end + start;
		if(size > remain)
			size = remain;
		if(end + size >= Size)
		{
			int l = Size - end;
			memcpy(buf + end, data, l);
			memcpy(buf, data + l, size - l);
			end = size - l;
		}
		else
		{
			memcpy(buf + end, data, size);
			end += size;
		}
		return size;
	}

	int Read(char* data, int size)
	{
		int len = end - start;
		if(len < 0)
			len += Size;
		if(len < size)
			size = len;
		if(start + size >= Size)
		{
			int l = Size - start;
			memcpy(data, buf + start, l);
			memcpy(data + l, buf, size - l);
			start = size - l;
		}
		else
		{
			memcpy(data, buf + start, size);
			start += size;
		}
		return size;
	}

	int Length()
	{
		int len = end - start;
		if(len < 0)
			len += Size;
		return len;
	}
};

CComCriticalSection lock;
static MMRESULT (WINAPI * fnwaveOutWrite)(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) = waveOutWrite;
static DWORD (WINAPI * fntimeGetTime)(void) = timeGetTime;

static MMRESULT (WINAPI* fnwaveOutOpen)( __out_opt LPHWAVEOUT phwo, __in UINT uDeviceID,
    __in LPCWAVEFORMATEX pwfx, __in_opt DWORD_PTR dwCallback, __in_opt DWORD_PTR dwInstance, __in DWORD fdwOpen) = waveOutOpen;


HWAVEOUT outHandle = 0;
WAVEFORMATEX WaveFormatHead;
RingBuffer<256*1024> SoundBuffer;

MMRESULT WINAPI waveOutOpenHooked( __out_opt LPHWAVEOUT phwo, __in UINT uDeviceID,
    __in LPCWAVEFORMATEX pwfx, __in_opt DWORD_PTR dwCallback, __in_opt DWORD_PTR dwInstance, __in DWORD fdwOpen)
{
	MMRESULT ret = fnwaveOutOpen(phwo, uDeviceID, pwfx, dwCallback, dwInstance, fdwOpen);
	if((fdwOpen & WAVE_FORMAT_QUERY) == 0 && uDeviceID == -1)
	{
		printf("waveOutOpen %x\n", *phwo);
		printf("channels:%d, nSamplesPerSec:%d, wBitsPerSample:%d\n", pwfx->nChannels, pwfx->nSamplesPerSec, pwfx->wBitsPerSample);
		WaveFormatHead = *pwfx;
		outHandle = *phwo;
	}
	return ret;
}


MMRESULT WINAPI hookedwaveOutWrite(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh)
{
	/*
	char* flags[] = {
		"DONE ", 
		"PREPARED ", 
		"BEGINLOOP ", 
		"ENDLOOP ", 
		"INQUEUE "
	};
	std::string flag;
	for(int i = 0; i < 5; ++i)
		if((pwh->dwFlags & (1 << i)) != 0)
			flag += flags[i];
*/
	//static int pos = 0;
	//printf("waveOutWrite %d msec.\n", pos);
	//pos += 1000 * pwh->dwBufferLength / (WaveFormatHead.nSamplesPerSec * (WaveFormatHead.wBitsPerSample / 8) * WaveFormatHead .nChannels);
	//printf("waveOutWrite %x %d %d %s\n", hwo, pwh->dwBufferLength, flag.c_str());
	//return MMSYSERR_NOERROR;
	if(outHandle == hwo && pwh->dwFlags == WHDR_PREPARED)
	{
		lock.Lock();
		SoundBuffer.Write(pwh->lpData, pwh->dwBufferLength);
		lock.Unlock();
	}
	return fnwaveOutWrite(hwo, pwh, cbwh);
}

DWORD (WINAPI timeGetTimeHooked)(void)
{
	printf("timeGetTime\n");
	return 0;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID)
{
    LONG error;
	if (dwReason == DLL_PROCESS_ATTACH) {
		DetourRestoreAfterWith();

		printf("Starting.\n");

		DetourTransactionBegin();
		DetourUpdateThread(GetCurrentThread());
		DetourAttach(&(PVOID&)fnwaveOutOpen, waveOutOpenHooked);
		DetourAttach(&(PVOID&)fnwaveOutWrite, hookedwaveOutWrite);
		error = DetourTransactionCommit();

		if (error == NO_ERROR)
			printf("Detoured sound.\n");
		else
			printf("Error detouring : %d\n", error);
		lock.Init();
	}
	else if (dwReason == DLL_PROCESS_DETACH) {
		DetourTransactionBegin();
		DetourUpdateThread(GetCurrentThread());
		DetourDetach(&(PVOID&)fnwaveOutOpen, waveOutOpenHooked);
		DetourDetach(&(PVOID&)fnwaveOutWrite, hookedwaveOutWrite);
		error = DetourTransactionCommit();

		printf("Removed detour (result=%d)\n",  error);
	}

	return TRUE;
}


extern "C" {

__declspec(dllexport) void GetWaveFormat(WAVEFORMATEX* format)
{
	*format = WaveFormatHead;
}

__declspec(dllexport) int GetSoundData(void* buf, int size)
{
	lock.Lock();
	int len = SoundBuffer.Read((char*)buf, size);
	lock.Unlock();
	return len;
}

};
