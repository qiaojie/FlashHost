// SoundHook.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "SoundHook.h"
#include <mmsystem.h>

extern WAVEFORMATEX WaveFormatHead;

// This is an example of an exported function.
SOUNDHOOK_API int fnSoundHook(void)
{
	return 42;
}