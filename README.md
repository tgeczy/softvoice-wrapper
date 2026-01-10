# softvoice-wrapper

A small C/C++ wrapper DLL around the classic SoftVoice text-to-speech engine. The wrapper loads the
original SoftVoice runtime (for example, `tibase32.dll`), hooks the WinMM waveOut APIs with MinHook,
and exposes a simple pull-based audio queue so consumers (like the included NVDA driver) can fetch
PCM audio in manageable chunks.

## What this wrapper does

* Loads SoftVoice’s 32-bit engine DLLs and resolves the core export functions at runtime.
* Spins up a dedicated thread with a message window so the engine can post window messages as it
  expects.
* Hooks `waveOut*` calls to capture generated audio and place it into a queue.
* Exposes a small C ABI for:
  * Initialization and teardown (`sv_initW`, `sv_free`).
  * Starting/stopping speech (`sv_startSpeakW`, `sv_stop`).
  * Reading audio frames and event items (`sv_read`).
  * Getting/setting voice and tuning parameters (`sv_setRate`, `sv_setPitch`, `sv_setPersonality`,
    `sv_setVoicingMode`, etc.).

## How it works (high level)

1. **Initialization**: The wrapper loads `tibase32.dll` (and related engine DLLs) from a provided
   path, resolves the SoftVoice exports, and configures internal state.
2. **Audio capture**: The SoftVoice engine calls `waveOutOpen/Write/Close` to emit audio. The wrapper
   uses MinHook to intercept these calls, copies the audio into an internal queue, and applies
   backpressure so the engine doesn’t overrun the consumer.
3. **Consumer pull**: Client code (for example, the NVDA driver in `src/sv.py`) repeatedly calls
   `sv_read` to pull audio frames and event markers (done/error/index). This decouples synthesis from
   playback.

## SoftVoice DLLs are **not** included

This repository **does not** include `tibase32.dll` or any other SoftVoice binaries. No parts of
SoftVoice are redistributed here. You must provide your own licensed SoftVoice installation and
point the wrapper at the engine DLLs on your system.

## SoftVoice historical note (brief)

SoftVoice was a widely used TTS engine from the late 1990s/early 2000s era, commonly bundled with
accessibility software and utilities on Windows. It’s remembered for its distinctive voices and the
ability to run fully offline, which made it popular on the hardware of the time. This wrapper exists
to make that legacy engine usable in modern assistive tech pipelines without modifying the original
binaries.

## Build notes

* SoftVoice is 32-bit, so the wrapper must be built as **x86**.
* The project links against MinHook (included under `src/minhook`).
* Build with CMake using the included `CMakeLists.txt`.

## `softvoice-say` sample app

The repository now includes a small standalone Windows dialog app at
`src/softvoice-say.cpp`. It is a lightweight "Speak Window" that exercises the
wrapper DLL: enter or load text, tweak the same voice parameters exposed by the
NVDA driver, and either speak immediately or save the output to a 11025 Hz,
16-bit, mono WAV file. The source expects the dialog resource in
`src/softvoice-say.rc` (and `resource.h`).

Build notes:

* Build x86 (SoftVoice is 32-bit).
* Place `softvoice_wrapper.dll` and `tibase32.dll` (plus any SoftVoice language
  DLLs) next to the built EXE.

## Thanks

Huge thanks for checking this project out and for helping keep classic TTS engines accessible.
