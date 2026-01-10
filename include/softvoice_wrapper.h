// softvoice_wrapper.h
//
// SoftVoice (late 90s) wrapper DLL for NVDA.
//
// Design:
// - Hooks winmm waveOut* (via MinHook) to capture PCM written by tibase32.dll.
// - Exposes a pull API (sv_read) so the NVDA synth driver can feed nvwave.WavePlayer.
// - Runs SoftVoice calls on a dedicated worker thread with a message-only window.
// - Supports repeated init/free cycles within the same process (ctypes keeps DLLs loaded).

#pragma once

#include <stddef.h>
#include <stdint.h>

#ifdef _WIN32
  #ifdef SVWRAP_EXPORTS
    #define SV_API __declspec(dllexport)
  #else
    #define SV_API __declspec(dllimport)
  #endif
#else
  #define SV_API
#endif

#ifdef __cplusplus
extern "C" {
#endif

typedef struct SV_STATE SV_STATE;

enum SV_ITEM_TYPE {
  SV_ITEM_NONE  = 0,
  SV_ITEM_AUDIO = 1,
  SV_ITEM_DONE  = 2,
  SV_ITEM_ERROR = 3,
};

// Lifecycle
SV_API SV_STATE* __cdecl sv_initW(const wchar_t* baseDllPath, int initialVoice);
SV_API void __cdecl sv_free(SV_STATE* s);

// Stop / cancel any current or queued speech
SV_API void __cdecl sv_stop(SV_STATE* s);

// Queue text for synthesis (UTF-16). Returns 0 on success.
SV_API int __cdecl sv_startSpeakW(SV_STATE* s, const wchar_t* text);

// Pull items out of the wrapper.
// - outType: SV_ITEM_NONE, SV_ITEM_AUDIO, SV_ITEM_DONE, SV_ITEM_ERROR
// - outValue: optional integer payload (currently unused by the NVDA driver)
// Returns:
// - For audio: number of bytes copied into outBuf (<= outBufBytes)
// - For markers: 0
SV_API int __cdecl sv_read(SV_STATE* s, int* outType, int* outValue, uint8_t* outBuf, int outBufBytes);

// SoftVoice numeric settings (int ranges follow tibase32 semantics)
SV_API void __cdecl sv_setVoice(SV_STATE* s, int voice);
SV_API int __cdecl sv_getVoice(SV_STATE* s);

SV_API void __cdecl sv_setRate(SV_STATE* s, int rate);
SV_API int __cdecl sv_getRate(SV_STATE* s);

SV_API void __cdecl sv_setPitch(SV_STATE* s, int pitch);
SV_API int __cdecl sv_getPitch(SV_STATE* s);

SV_API void __cdecl sv_setF0Range(SV_STATE* s, int range);
SV_API int __cdecl sv_getF0Range(SV_STATE* s);

SV_API void __cdecl sv_setF0Perturb(SV_STATE* s, int perturb);
SV_API int __cdecl sv_getF0Perturb(SV_STATE* s);

SV_API void __cdecl sv_setVowelFactor(SV_STATE* s, int factor);
SV_API int __cdecl sv_getVowelFactor(SV_STATE* s);

SV_API void __cdecl sv_setAVBias(SV_STATE* s, int bias);
SV_API int __cdecl sv_getAVBias(SV_STATE* s);

SV_API void __cdecl sv_setAFBias(SV_STATE* s, int bias);
SV_API int __cdecl sv_getAFBias(SV_STATE* s);

SV_API void __cdecl sv_setAHBias(SV_STATE* s, int bias);
SV_API int __cdecl sv_getAHBias(SV_STATE* s);

// Personality / style settings
SV_API void __cdecl sv_setPersonality(SV_STATE* s, int p);
SV_API int __cdecl sv_getPersonality(SV_STATE* s);

SV_API void __cdecl sv_setF0Style(SV_STATE* s, int style);
SV_API int __cdecl sv_getF0Style(SV_STATE* s);

SV_API void __cdecl sv_setVoicingMode(SV_STATE* s, int mode);
SV_API int __cdecl sv_getVoicingMode(SV_STATE* s);

SV_API void __cdecl sv_setGender(SV_STATE* s, int g);
SV_API int __cdecl sv_getGender(SV_STATE* s);

SV_API void __cdecl sv_setGlottalSource(SV_STATE* s, int gs);
SV_API int __cdecl sv_getGlottalSource(SV_STATE* s);

SV_API void __cdecl sv_setSpeakingMode(SV_STATE* s, int sm);
SV_API int __cdecl sv_getSpeakingMode(SV_STATE* s);

// Wrapper-only tuning
// maxLeadMs controls how far ahead of realtime we allow SoftVoice to synthesize.
// Larger values reduce underflow but increase index latency.
SV_API void __cdecl sv_setMaxLeadMs(SV_STATE* s, int ms);
SV_API int __cdecl sv_getMaxLeadMs(SV_STATE* s);

// Enable/disable trimming of leading/trailing silence inside each synthesis chunk.
SV_API void __cdecl sv_setTrimSilence(SV_STATE* s, int enabled);
SV_API int __cdecl sv_getTrimSilence(SV_STATE* s);

// "Pause factor" controls how aggressive the wrapper trims silence.
// 0 disables trimming even if sv_setTrimSilence(...,1) is enabled.
// Typical range: 0..100.
SV_API void __cdecl sv_setPauseFactor(SV_STATE* s, int factor);
SV_API int __cdecl sv_getPauseFactor(SV_STATE* s);

// Output audio format for the current synthesis session.
SV_API int __cdecl sv_getFormat(SV_STATE* s, int* sampleRate, int* channels, int* bitsPerSample);

#ifdef __cplusplus
} // extern "C"
#endif
