// softvoice_wrapper.cpp
//
// SoftVoice (late 90s) wrapper DLL for NVDA.
// - Hooks winmm waveOut* to capture audio into a queue.
// - Exposes a simple pull API (sv_read) so the NVDA Python driver can feed NVWave.
// - Runs SoftVoice calls on a dedicated thread with a message window (SoftVoice uses window messages).
//
// Key fixes vs early versions:
// - Apply settings in a safe order, and only when needed:
//   * Personality (variant) is treated like a preset: apply it first, then re-apply user numeric params.
//   * Optional "style" params (voicing mode, glottal source, etc.) are only applied if explicitly set,
//     so personalities like Robot/Martian can keep their internal presets.
// - Sprint-and-wait buffering: the waveOutWrite hook allows SoftVoice to sprint until the queue
//   fills, then waits in chunk-sized increments to apply backpressure.
// - Optional conservative silence trimming to reduce chunk-boundary pauses.
//
// NOTE: SoftVoice is 32-bit (tibase32.dll etc). Build this wrapper as 32-bit.

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#ifndef SV_API
#define SV_API __declspec(dllexport)
#endif

// Stream item types (match the NVDA driver: sv.py).
#ifndef SV_ITEM_NONE
#define SV_ITEM_NONE 0
#endif
#ifndef SV_ITEM_AUDIO
#define SV_ITEM_AUDIO 1
#endif
#ifndef SV_ITEM_DONE
#define SV_ITEM_DONE 2
#endif
#ifndef SV_ITEM_ERROR
#define SV_ITEM_ERROR 3
#endif

#include <windows.h>
#include <mmsystem.h>
#include <intrin.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <mutex>
#include <string>
#include <thread>
#include <vector>
#include <chrono>
#include <climits>

#include "MinHook.h"

#pragma comment(lib, "user32.lib")

// ------------------------------------------------------------
// SoftVoice exports (tibase32.dll)
// ------------------------------------------------------------
// Exports in the real DLL are stdcall-decorated in 32-bit builds (e.g. "_SVTTS@32").
// We try both undecorated and decorated names in GetProcAddress for robustness.

typedef int(__stdcall* SVOpenSpeechFunc)(int* outHandle, HWND hwnd, int msg, int voice, int flags);
typedef int(__stdcall* SVCloseSpeechFunc)(int handle);
typedef int(__stdcall* SVAbortFunc)(int handle);

// Many setters are (handle, int) and usually return int (0=ok).
typedef int(__stdcall* SVSet2IntFunc)(int handle, int val);

typedef int(__stdcall* SVTTSFunc)(
    int handle,
    const char* text,
    int a,
    int b,
    HWND hwnd,
    int c,
    int d,
    int e
);

// ------------------------------------------------------------
// SEH-safe helpers (NO fancy templates). Keep these tiny.
// ------------------------------------------------------------
static __declspec(noinline) int seh_svOpenSpeech(SVOpenSpeechFunc fn, int* outHandle, HWND hwnd, int msg, int voice, int flags) {
    if (!fn || !outHandle) return -1;
    __try { return fn(outHandle, hwnd, msg, voice, flags); }
    __except (EXCEPTION_EXECUTE_HANDLER) { return -1; }
}

static __declspec(noinline) int seh_svCloseSpeech(SVCloseSpeechFunc fn, int handle) {
    if (!fn) return -1;
    __try { return fn(handle); }
    __except (EXCEPTION_EXECUTE_HANDLER) { return -1; }
}

static __declspec(noinline) int seh_svAbort(SVAbortFunc fn, int handle) {
    if (!fn) return -1;
    __try { return fn(handle); }
    __except (EXCEPTION_EXECUTE_HANDLER) { return -1; }
}

static __declspec(noinline) int seh_svSet2Int(SVSet2IntFunc fn, int handle, int v) {
    if (!fn) return -1;
    __try { return fn(handle, v); }
    __except (EXCEPTION_EXECUTE_HANDLER) { return -1; }
}

static __declspec(noinline) int seh_svTTS(SVTTSFunc fn, int handle, const char* text, int a, int b, HWND hwnd, int c, int d, int e) {
    if (!fn || !text) return -1;
    __try { return fn(handle, text, a, b, hwnd, c, d, e); }
    __except (EXCEPTION_EXECUTE_HANDLER) { return -1; }
}

// ------------------------------------------------------------
// WinMM function pointer types + originals
// ------------------------------------------------------------
typedef MMRESULT(WINAPI* waveOutOpenFunc)(LPHWAVEOUT, UINT, LPCWAVEFORMATEX, DWORD_PTR, DWORD_PTR, DWORD);
typedef MMRESULT(WINAPI* waveOutPrepareHeaderFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutWriteFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutUnprepareHeaderFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutResetFunc)(HWAVEOUT);
typedef MMRESULT(WINAPI* waveOutCloseFunc)(HWAVEOUT);

static waveOutOpenFunc            g_waveOutOpenOrig = nullptr;
static waveOutPrepareHeaderFunc   g_waveOutPrepareHeaderOrig = nullptr;
static waveOutWriteFunc           g_waveOutWriteOrig = nullptr;
static waveOutUnprepareHeaderFunc g_waveOutUnprepareHeaderOrig = nullptr;
static waveOutResetFunc           g_waveOutResetOrig = nullptr;
static waveOutCloseFunc           g_waveOutCloseOrig = nullptr;

// ------------------------------------------------------------
// Stream queue items
// ------------------------------------------------------------
struct StreamItem {
    int type = SV_ITEM_NONE;
    int value = 0;
    uint32_t gen = 0;
    std::vector<uint8_t> data;
    size_t offset = 0;

    StreamItem() = default;
    StreamItem(int t, int v, uint32_t g) : type(t), value(v), gen(g), offset(0) {}
};

// ------------------------------------------------------------
// Command queue
// ------------------------------------------------------------
struct Cmd {
    enum Type { CMD_SPEAK, CMD_QUIT } type = CMD_SPEAK;
    uint32_t cancelSnapshot = 0;
    std::wstring text;
};

// ------------------------------------------------------------
// Settings state (small helper)
// ------------------------------------------------------------
struct SettingInt {
    std::atomic<int> value{ 0 };
    std::atomic<int> dirty{ 0 };    // 1 if needs applying
    std::atomic<int> userSet{ 0 };  // 1 if ever set by Python (for optional style params)
};

// ------------------------------------------------------------
// Global wrapper state
// ------------------------------------------------------------
struct SV_STATE {
    // DLLs
    HMODULE baseModule = nullptr;
    HMODULE engModule = nullptr;
    HMODULE spanModule = nullptr;

    // Path we were initialized with (used to validate repeated inits).
    std::wstring baseDllPath;

    // SoftVoice sync message routing.
    // We learn the actual message id used by the engine to avoid false DONE events
    // from unrelated Win32 messages like WM_TIMER.
    UINT svSyncMsg = 0;       // RegisterWindowMessageW(L"SVSyncMessages") (0 if unavailable)
    UINT activeSyncMsg = 0;   // Message id actually observed from the engine (learned)

    // Exports
    SVOpenSpeechFunc  svOpenSpeech = nullptr;
    SVCloseSpeechFunc svCloseSpeech = nullptr;
    SVAbortFunc       svAbort = nullptr;
    SVTTSFunc         svTTS = nullptr;

    // Optional: language switch without reopening.
    SVSet2IntFunc svSetLanguage = nullptr;

    // Setters
    SVSet2IntFunc svSetRate = nullptr;
    SVSet2IntFunc svSetPitch = nullptr;
    SVSet2IntFunc svSetF0Range = nullptr;
    SVSet2IntFunc svSetF0Perturb = nullptr;
    SVSet2IntFunc svSetVowelFactor = nullptr;

    SVSet2IntFunc svSetAVBias = nullptr;
    SVSet2IntFunc svSetAFBias = nullptr;
    SVSet2IntFunc svSetAHBias = nullptr;

    SVSet2IntFunc svSetPersonality = nullptr;
    SVSet2IntFunc svSetF0Style = nullptr;
    SVSet2IntFunc svSetVoicingMode = nullptr;
    SVSet2IntFunc svSetGender = nullptr;
    SVSet2IntFunc svSetGlottalSource = nullptr;
    SVSet2IntFunc svSetSpeakingMode = nullptr;

    // SoftVoice handle (opened/used only on worker thread)
    int svHandle = 0;
    int currentVoice = 1;

    // Message window (owned by worker thread)
    HWND msgWnd = nullptr;

    // waveOutOpen capture
    WAVEFORMATEX lastFormat = {};
    bool formatValid = false;

    DWORD callbackType = 0;
    DWORD_PTR callbackTarget = 0;
    DWORD_PTR callbackInstance = 0;

    // Events
    HANDLE startEvent = nullptr; // signaled by WndProc when SoftVoice starts (best effort)
    HANDLE doneEvent = nullptr;  // signaled by WndProc when SoftVoice finishes
    HANDLE stopEvent = nullptr;  // signaled by sv_stop
    HANDLE cmdEvent = nullptr;   // signaled when commands are queued
    HANDLE initEvent = nullptr;  // signaled by worker when initial init is complete
    std::atomic<int> initOk{ 0 }; // 1=ok, -1=failed

    // Cancel + generations
    std::atomic<uint32_t> cancelToken{ 1 };
    std::atomic<uint32_t> genCounter{ 1 };
    std::atomic<uint32_t> activeGen{ 0 };   // hooks capture only while nonzero
    std::atomic<uint32_t> currentGen{ 0 };  // reader consumes only this gen

    // Output pacing data
    std::atomic<uint64_t> bytesPerSec{ 0 };
    std::atomic<ULONGLONG> lastAudioTick{ 0 };

    // Legacy: allow SoftVoice to synthesize ahead of a virtual playback clock.
    // Kept for compatibility with older builds, but currently unused.
    std::atomic<int> maxLeadMs{ 2000 };
    std::atomic<int> autoLead{ 1 }; // if 1, wrapper will tweak maxLeadMs when speaking mode changes

    // Optional silence trim
    std::atomic<int> trimSilence{ 1 };
    // Pause factor (0..100). Higher values trim more silence at chunk boundaries.
    std::atomic<int> pauseFactor{ 50 };
    std::atomic<uint32_t> leadTrimDoneGen{ 0 };
    std::atomic<uint32_t> tailTrimDoneGen{ 0 };

    // Desired settings (setters store these; worker applies them before SVTTS)
    // Numeric params are always considered user-set (NVDA will show them as controls).
    SettingInt rate;
    SettingInt pitch;
    SettingInt f0Range;
    SettingInt f0Perturb;
    SettingInt vowelFactor;

    SettingInt avBias;
    SettingInt afBias;
    SettingInt ahBias;

    // Personality is a preset; apply only if user set it (variant control).
    SettingInt personality;

    // Style params: apply only if user set them, so personalities can keep their own presets by default.
    SettingInt f0Style;
    SettingInt voicingMode;
    SettingInt gender;
    SettingInt glottalSource;
    SettingInt speakingMode;

    // Voice/language (always user-set)
    SettingInt voice;

    // Command queue
    std::mutex cmdMtx;
    std::deque<Cmd> cmdQ;
    std::thread worker;

    // Output queue
    std::mutex outMtx;
    std::deque<StreamItem> outQ;
    size_t queuedAudioBytes = 0;

    size_t maxBufferedBytes = 0;
    size_t maxQueueItems = 8192;
};

static SV_STATE* g_state = nullptr;
// NVDA can briefly construct a new SynthDriver instance before terminating the old one.
// The wrapper is a singleton (SoftVoice is not designed for multi-instance use),
// so we keep a small refcount and allow sv_initW to return the existing instance.
static std::mutex g_globalMtx;
static std::atomic<long> g_refCount{ 0 };

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------
static FARPROC getProcMaybeDecorated(HMODULE mod, const char* undecorated, const char* decorated) {
    if (!mod) return nullptr;
    FARPROC p = nullptr;
    if (undecorated) p = GetProcAddress(mod, undecorated);
    if (!p && decorated) p = GetProcAddress(mod, decorated);
    return p;
}

static bool isCallerFromSoftVoice(SV_STATE* s) {
    if (!s) return false;
    void* ra = _ReturnAddress();
    HMODULE caller = nullptr;
    if (!GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(ra),
        &caller
    )) {
        return false;
    }
    return caller && (caller == s->baseModule || caller == s->engModule || caller == s->spanModule);
}

static void signalWaveOutMessage(SV_STATE* s, UINT msg, WAVEHDR* hdr) {
    if (!s) return;

    const DWORD cbType = (s->callbackType & CALLBACK_TYPEMASK);

    if (cbType == CALLBACK_FUNCTION) {
        typedef void(CALLBACK* WaveOutProc)(HWAVEOUT, UINT, DWORD_PTR, DWORD_PTR, DWORD_PTR);
        WaveOutProc proc = reinterpret_cast<WaveOutProc>(s->callbackTarget);
        if (proc) {
            proc(reinterpret_cast<HWAVEOUT>(s), msg, s->callbackInstance,
                reinterpret_cast<DWORD_PTR>(hdr), 0);
        }
        return;
    }

    if (cbType == CALLBACK_WINDOW) {
        HWND hwnd = reinterpret_cast<HWND>(s->callbackTarget);
        if (!hwnd) return;
        UINT mmMsg = 0;
        if (msg == WOM_OPEN) mmMsg = MM_WOM_OPEN;
        else if (msg == WOM_CLOSE) mmMsg = MM_WOM_CLOSE;
        else if (msg == WOM_DONE) mmMsg = MM_WOM_DONE;
        if (mmMsg) PostMessageW(hwnd, mmMsg, reinterpret_cast<WPARAM>(s), reinterpret_cast<LPARAM>(hdr));
        return;
    }

    if (cbType == CALLBACK_THREAD) {
        DWORD tid = static_cast<DWORD>(s->callbackTarget);
        UINT mmMsg = 0;
        if (msg == WOM_OPEN) mmMsg = MM_WOM_OPEN;
        else if (msg == WOM_CLOSE) mmMsg = MM_WOM_CLOSE;
        else if (msg == WOM_DONE) mmMsg = MM_WOM_DONE;
        if (mmMsg && tid) PostThreadMessageW(tid, mmMsg, reinterpret_cast<WPARAM>(s), reinterpret_cast<LPARAM>(hdr));
        return;
    }

    if (cbType == CALLBACK_EVENT) {
        HANDLE ev = reinterpret_cast<HANDLE>(s->callbackTarget);
        if (ev) SetEvent(ev);
        return;
    }
}

static void clearOutputQueueLocked(SV_STATE* s) {
    s->outQ.clear();
    s->queuedAudioBytes = 0;
}

static void computeBufferLimits(SV_STATE* s) {
    // SoftVoice audio is small; allow enough buffering to never drop during normal speech.
    uint64_t bps = s->bytesPerSec.load(std::memory_order_relaxed);
    if (bps == 0) bps = 22050;

    // 60 seconds max buffer.
    uint64_t bytes = bps * 60ULL;

    if (bytes < 256ULL * 1024ULL) bytes = 256ULL * 1024ULL;
    if (bytes > 8ULL * 1024ULL * 1024ULL) bytes = 8ULL * 1024ULL * 1024ULL;

    s->maxBufferedBytes = (size_t)bytes;
}

static std::string sanitizeForSoftVoiceCp1252(const wchar_t* in) {
    if (!in) return std::string();

    // Basic cleanup: strip control chars, collapse whitespace.
    std::wstring tmp;
    tmp.reserve(wcslen(in));

    bool prevSpace = true;
    for (const wchar_t* p = in; *p; ++p) {
        wchar_t ch = *p;

        if (ch == 0x00A0) ch = L' '; // NBSP

        // Replace most control chars with space.
        if ((ch < 0x20 && ch != L'\r' && ch != L'\n' && ch != L'\t') || (ch >= 0x7F && ch <= 0x9F)) {
            ch = L' ';
        }

        const bool isSpace = (ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n');
        if (isSpace) {
            if (!prevSpace) tmp.push_back(L' ');
            prevSpace = true;
        } else {
            tmp.push_back(ch);
            prevSpace = false;
        }
    }
    while (!tmp.empty() && tmp.back() == L' ') tmp.pop_back();

    if (tmp.empty()) return std::string();

    BOOL usedDefault = FALSE;
    int blen = WideCharToMultiByte(
        1252,
        WC_NO_BEST_FIT_CHARS,
        tmp.c_str(),
        -1,
        nullptr,
        0,
        " ",
        &usedDefault
    );
    if (blen <= 0) return std::string();

    std::string out;
    out.resize((size_t)blen);
    WideCharToMultiByte(
        1252,
        WC_NO_BEST_FIT_CHARS,
        tmp.c_str(),
        -1,
        &out[0],
        blen,
        " ",
        &usedDefault
    );

    // Drop trailing NUL.
    if (!out.empty() && out.back() == '\0') out.pop_back();

    // SoftVoice tends to behave better with spaces than '?' placeholders.
    for (char& c : out) {
        if (c == '?') c = ' ';
    }

    // Collapse spaces again (after replacements).
    std::string collapsed;
    collapsed.reserve(out.size());
    prevSpace = true;
    for (char c : out) {
        const bool isSpace = (c == ' ' || c == '\t' || c == '\r' || c == '\n');
        if (isSpace) {
            if (!prevSpace) collapsed.push_back(' ');
            prevSpace = true;
        } else {
            collapsed.push_back(c);
            prevSpace = false;
        }
    }
    while (!collapsed.empty() && collapsed.back() == ' ') collapsed.pop_back();
    return collapsed;
}

static std::vector<std::string> splitSoftVoiceTextIntoChunks(const std::string& text, size_t chunkChars) {
    std::vector<std::string> out;
    if (text.empty() || chunkChars == 0) return out;

    size_t start = 0;
    while (start < text.size()) {
        const size_t remaining = text.size() - start;
        if (remaining <= chunkChars) {
            out.push_back(text.substr(start));
            break;
        }

        // Find the first space after the chunk boundary so we don't split mid-word.
        size_t split = text.find(' ', start + chunkChars);
        if (split == std::string::npos) {
            // No space found; hard-split to guarantee progress.
            split = start + chunkChars;
        }

        if (split > start) {
            out.push_back(text.substr(start, split - start));
        }

        // Skip one (or more) spaces so the next chunk doesn't start with whitespace.
        start = split;
        while (start < text.size() && text[start] == ' ') ++start;
    }

    return out;
}

static void enqueueAudioFromHook(SV_STATE* s, uint32_t gen, const void* data, size_t size, bool* outWasEmpty, bool* outWasFull) {
    if (outWasEmpty) *outWasEmpty = false;
    if (outWasFull) *outWasFull = false;
    if (!s || !data || size == 0) return;

    std::vector<uint8_t> copied(size);
    std::memcpy(copied.data(), data, size);

    s->lastAudioTick.store(GetTickCount64(), std::memory_order_relaxed);

    std::lock_guard<std::mutex> g(s->outMtx);

    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0 || gen != curGen) return;

    const size_t limit = (s->maxBufferedBytes > 0) ? s->maxBufferedBytes : (size_t)(512 * 1024);
    const bool wasEmpty = (s->queuedAudioBytes == 0);
    const bool wasFull = (s->queuedAudioBytes >= limit);
    if (outWasEmpty) *outWasEmpty = wasEmpty;
    if (outWasFull) *outWasFull = wasFull;

    auto dropOneAudio = [&]() -> bool {
        for (auto it = s->outQ.begin(); it != s->outQ.end(); ++it) {
            if (it->type == SV_ITEM_AUDIO) {
                size_t remaining = 0;
                if (it->data.size() > it->offset) remaining = it->data.size() - it->offset;
                if (s->queuedAudioBytes >= remaining) s->queuedAudioBytes -= remaining;
                else s->queuedAudioBytes = 0;
                s->outQ.erase(it);
                return true;
            }
        }
        return false;
    };

    while ((s->queuedAudioBytes + copied.size() > limit) || (s->outQ.size() >= s->maxQueueItems)) {
        if (!dropOneAudio()) return;
    }

    StreamItem it(SV_ITEM_AUDIO, 0, gen);
    it.data = std::move(copied);
    it.offset = 0;

    s->queuedAudioBytes += it.data.size();
    s->outQ.push_back(std::move(it));
}

static void pushMarker(SV_STATE* s, int type, int value, uint32_t gen) {
    std::lock_guard<std::mutex> g(s->outMtx);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0 || gen != curGen) return;
    s->outQ.push_back(StreamItem(type, value, gen));
}

// ------------------------------------------------------------
// Hooks
// ------------------------------------------------------------
static MMRESULT WINAPI hook_waveOutOpen(
    LPHWAVEOUT phwo,
    UINT uDeviceID,
    LPCWAVEFORMATEX pwfx,
    DWORD_PTR dwCallback,
    DWORD_PTR dwInstance,
    DWORD fdwOpen
) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutOpenOrig ? g_waveOutOpenOrig(phwo, uDeviceID, pwfx, dwCallback, dwInstance, fdwOpen)
            : MMSYSERR_ERROR;
    }

    if (phwo) *phwo = reinterpret_cast<HWAVEOUT>(s);

    if (pwfx) {
        s->lastFormat = *pwfx;
        s->formatValid = true;

        uint64_t bps = 0;
        if (pwfx->nAvgBytesPerSec) bps = (uint64_t)pwfx->nAvgBytesPerSec;
        if (!bps && pwfx->nSamplesPerSec && pwfx->nBlockAlign) bps = (uint64_t)pwfx->nSamplesPerSec * (uint64_t)pwfx->nBlockAlign;
        if (!bps) bps = 22050;
        s->bytesPerSec.store(bps, std::memory_order_relaxed);
        computeBufferLimits(s);
    }

    s->callbackType = fdwOpen;
    s->callbackTarget = dwCallback;
    s->callbackInstance = dwInstance;

    signalWaveOutMessage(s, WOM_OPEN, nullptr);
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutPrepareHeader(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutPrepareHeaderOrig ? g_waveOutPrepareHeaderOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }
    if (pwh) pwh->dwFlags |= WHDR_PREPARED;
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutUnprepareHeader(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutUnprepareHeaderOrig ? g_waveOutUnprepareHeaderOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }
    if (pwh) pwh->dwFlags &= ~WHDR_PREPARED;
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutWrite(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutWriteOrig ? g_waveOutWriteOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }

    if (!pwh) return MMSYSERR_INVALPARAM;

    const uint32_t gen = s->activeGen.load(std::memory_order_relaxed);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    const bool capturing = (gen != 0 && gen == curGen);

    bool bufferWasEmpty = false;
    bool bufferWasFull = false;
    if (capturing && pwh->lpData && pwh->dwBufferLength > 0) {
        enqueueAudioFromHook(s, gen, pwh->lpData, (size_t)pwh->dwBufferLength, &bufferWasEmpty, &bufferWasFull);
    }

    // If we are not capturing (e.g. canceled), finish immediately.
    if (!capturing) {
        pwh->dwFlags |= WHDR_DONE;
        signalWaveOutMessage(s, WOM_DONE, pwh);
        return MMSYSERR_NOERROR;
    }

    if (!bufferWasEmpty && bufferWasFull && pwh->dwBufferLength > 0) {
        uint64_t bps = s->bytesPerSec.load(std::memory_order_relaxed);
        if (bps == 0) bps = 22050;
        const uint64_t durMs64 = (bps ? (uint64_t)pwh->dwBufferLength * 1000ULL / bps : 0ULL);
        ULONGLONG sleepMs = (durMs64 > 0xFFFFFFFFULL) ? 0xFFFFFFFFu : (ULONGLONG)durMs64;

        // Sleep in small chunks; wake immediately on stop/cancel.
        while (sleepMs > 0) {
            if (s->activeGen.load(std::memory_order_relaxed) != curGen) break;
            DWORD chunk = (sleepMs > 5) ? 5 : (DWORD)sleepMs;
            DWORD w = WaitForSingleObject(s->stopEvent, chunk);
            if (w == WAIT_OBJECT_0) break;
            if (sleepMs >= chunk) sleepMs -= chunk;
            else sleepMs = 0;
        }
    }

    pwh->dwFlags |= WHDR_DONE;
    signalWaveOutMessage(s, WOM_DONE, pwh);
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutReset(HWAVEOUT hwo) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutResetOrig ? g_waveOutResetOrig(hwo) : MMSYSERR_ERROR;
    }
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutClose(HWAVEOUT hwo) {
    SV_STATE* s = g_state;
    if (!s || !isCallerFromSoftVoice(s)) {
        return g_waveOutCloseOrig ? g_waveOutCloseOrig(hwo) : MMSYSERR_ERROR;
    }
    signalWaveOutMessage(s, WOM_CLOSE, nullptr);
    return MMSYSERR_NOERROR;
}

static void ensureHooksInstalled() {
    static LONG once = 0;
    if (InterlockedCompareExchange(&once, 1, 0) != 0) return;

    // Make sure the modules are loaded before we try to hook them.
    // (MinHook's MH_CreateHookApi uses GetModuleHandle internally.)
    LoadLibraryW(L"winmm.dll");
    LoadLibraryW(L"winmmbase.dll"); // present on newer Windows; harmless if absent.

    MH_STATUS st = MH_Initialize();
    if (st != MH_OK && st != MH_ERROR_ALREADY_INITIALIZED) return;

    auto tryHookApi = [&](const wchar_t* mod, const char* proc, LPVOID detour, LPVOID* orig) -> bool {
        MH_STATUS rc = MH_CreateHookApi(mod, proc, detour, orig);
        return (rc == MH_OK || rc == MH_ERROR_ALREADY_CREATED);
    };

    auto hookEither = [&](const char* proc, LPVOID detour, LPVOID* orig) -> bool {
        // Try winmm.dll first, then fall back to winmmbase.dll (needed on some Windows builds).
        if (tryHookApi(L"winmm.dll", proc, detour, orig)) return true;
        if (tryHookApi(L"winmmbase.dll", proc, detour, orig)) return true;
        return false;
    };

    const bool okOpen = hookEither("waveOutOpen", (LPVOID)hook_waveOutOpen, (LPVOID*)&g_waveOutOpenOrig);
    const bool okPrep = hookEither("waveOutPrepareHeader", (LPVOID)hook_waveOutPrepareHeader, (LPVOID*)&g_waveOutPrepareHeaderOrig);
    const bool okUnprep = hookEither("waveOutUnprepareHeader", (LPVOID)hook_waveOutUnprepareHeader, (LPVOID*)&g_waveOutUnprepareHeaderOrig);
    const bool okWrite = hookEither("waveOutWrite", (LPVOID)hook_waveOutWrite, (LPVOID*)&g_waveOutWriteOrig);
    const bool okReset = hookEither("waveOutReset", (LPVOID)hook_waveOutReset, (LPVOID*)&g_waveOutResetOrig);
    const bool okClose = hookEither("waveOutClose", (LPVOID)hook_waveOutClose, (LPVOID*)&g_waveOutCloseOrig);

    // Avoid enabling partial hooks (that can lead to "silent" output).
    if (!(okOpen && okPrep && okUnprep && okWrite && okReset && okClose)) {
        return;
    }

    MH_EnableHook(MH_ALL_HOOKS);
}

// ------------------------------------------------------------
// SoftVoice message window
// ------------------------------------------------------------
static const wchar_t* kSvWrapWndClass = L"NVDA_SoftVoice_WrapWnd";

static LRESULT CALLBACK svWrapWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    SV_STATE* s = g_state;
    if (!s || hwnd != s->msgWnd) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    // SoftVoice uses small integer codes in wParam:
    // 1000 = started, 1001 = done, 1002 = error/other.
    // IMPORTANT: Do NOT treat these as events unless they arrive on the synthesizer's
    // dedicated sync message id. Otherwise, unrelated Win32 messages (notably WM_TIMER)
    // can carry the same wParam values and cause premature DONE, truncating speech.
    if (wParam != 1000 && wParam != 1001 && wParam != 1002) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    // If we've already learned the engine's sync message id, require it.
    if (s->activeSyncMsg != 0) {
        if ((UINT)msg != s->activeSyncMsg) {
            return DefWindowProcW(hwnd, msg, wParam, lParam);
        }
    } else {
        // Prefer the registered message id if available.
        if (s->svSyncMsg != 0 && (UINT)msg == s->svSyncMsg) {
            s->activeSyncMsg = (UINT)msg;
        } else {
            // Learn from the first plausible message id in the WM_USER/registered range.
            // This avoids WM_TIMER/WM_COMMAND (which are < WM_USER) collisions.
            if (msg < WM_USER) {
                return DefWindowProcW(hwnd, msg, wParam, lParam);
            }
            s->activeSyncMsg = (UINT)msg;
        }
    }

    if (wParam == 1000) {
        if (s->startEvent) SetEvent(s->startEvent);
        return 0;
    }
    if (wParam == 1001) {
        if (s->doneEvent) SetEvent(s->doneEvent);
        return 0;
    }
    // 1002: Treat as done; the worker will push an ERROR marker if needed.
    if (s->doneEvent) SetEvent(s->doneEvent);
    return 0;
}

static bool ensureMsgWindowCreated(SV_STATE* s) {
    if (!s) return false;

    // Register window class once.
    static LONG once = 0;
    if (InterlockedCompareExchange(&once, 1, 0) == 0) {
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(wc);
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.lpszClassName = kSvWrapWndClass;
        wc.lpfnWndProc = svWrapWndProc;
        if (!RegisterClassExW(&wc)) {
            DWORD e = GetLastError();
            if (e != ERROR_CLASS_ALREADY_EXISTS) return false;
        }
    }

    // Message-only window.
    HWND hwnd = CreateWindowExW(
        0,
        kSvWrapWndClass,
        L"",
        0,
        0, 0, 0, 0,
        HWND_MESSAGE,
        nullptr,
        GetModuleHandleW(nullptr),
        nullptr
    );
    if (!hwnd) return false;

    s->msgWnd = hwnd;
    return true;
}

// ------------------------------------------------------------
// Worker helpers
// ------------------------------------------------------------
static bool openVoiceOnWorker(SV_STATE* s, int voice) {
    if (!s || !s->svOpenSpeech) return false;

    // Close old.
    if (s->svHandle != 0 && s->svCloseSpeech) {
        seh_svCloseSpeech(s->svCloseSpeech, s->svHandle);
        s->svHandle = 0;
    }

    s->currentVoice = voice;
    int h = 0;

    // NOTE: SoftVoice's "msg" param isn't well documented. Passing 0 works for our message-only wnd;
    // the engine still seems to post its wParam status codes to that window.
    int rc = seh_svOpenSpeech(s->svOpenSpeech, &h, s->msgWnd, 0, voice, 0);
    if (rc != 0 || h == 0) {
        s->svHandle = 0;
        return false;
    }
    s->svHandle = h;
    return true;
}

static bool setLanguageOnWorker(SV_STATE* s, int voice) {
    if (!s || s->svHandle == 0 || !s->svSetLanguage) return false;
    int rc = seh_svSet2Int(s->svSetLanguage, s->svHandle, voice);
    if (rc == 0) {
        s->currentVoice = voice;
        return true;
    }
    return false;
}

static void applyNumericSettingsOnWorker(SV_STATE* s, bool force) {
    if (!s || s->svHandle == 0) return;
    const int h = s->svHandle;

    auto apply = [&](SettingInt& st, SVSet2IntFunc fn) {
        const int doApply = force || (st.dirty.exchange(0, std::memory_order_relaxed) != 0);
        if (doApply) {
            const int v = st.value.load(std::memory_order_relaxed);
            seh_svSet2Int(fn, h, v);
        }
    };

    apply(s->rate, s->svSetRate);
    apply(s->pitch, s->svSetPitch);
    apply(s->f0Range, s->svSetF0Range);
    apply(s->f0Perturb, s->svSetF0Perturb);
    apply(s->vowelFactor, s->svSetVowelFactor);

    apply(s->avBias, s->svSetAVBias);
    apply(s->afBias, s->svSetAFBias);
    apply(s->ahBias, s->svSetAHBias);
}

static void discardTimbreDirtyOnWorker(SV_STATE* s) {
    if (!s) return;
    s->pitch.dirty.store(0, std::memory_order_relaxed);
    s->f0Range.dirty.store(0, std::memory_order_relaxed);
    s->f0Perturb.dirty.store(0, std::memory_order_relaxed);
    s->vowelFactor.dirty.store(0, std::memory_order_relaxed);
    s->avBias.dirty.store(0, std::memory_order_relaxed);
    s->afBias.dirty.store(0, std::memory_order_relaxed);
    s->ahBias.dirty.store(0, std::memory_order_relaxed);
}

static void applyStyleSettingOnWorker(SV_STATE* s, SettingInt& st, SVSet2IntFunc fn, bool forceIfUserSet) {
    if (!s || s->svHandle == 0) return;
    if (!fn) return;

    const int userSet = st.userSet.load(std::memory_order_relaxed);
    if (!userSet) return;

    const bool doApply = forceIfUserSet || (st.dirty.exchange(0, std::memory_order_relaxed) != 0);
    if (!doApply) return;

    const int h = s->svHandle;
    const int v = st.value.load(std::memory_order_relaxed);
    seh_svSet2Int(fn, h, v);
}

static bool applyPersonalityOnWorker(SV_STATE* s, bool forceIfUserSet) {
    if (!s || s->svHandle == 0 || !s->svSetPersonality) return false;

    const int userSet = s->personality.userSet.load(std::memory_order_relaxed);
    if (!userSet) {
        // Clear dirty if any; we don't apply unless user explicitly used the control.
        s->personality.dirty.store(0, std::memory_order_relaxed);
        return false;
    }

    const bool doApply = forceIfUserSet || (s->personality.dirty.exchange(0, std::memory_order_relaxed) != 0);
    if (!doApply) return false;

    const int v = s->personality.value.load(std::memory_order_relaxed);
    seh_svSet2Int(s->svSetPersonality, s->svHandle, v);
    return true;
}

// ------------------------------------------------------------
// Silence trimming (conservative) — applied at read-time under outMtx.
// Supports PCM 8-bit unsigned and PCM 16-bit signed.
// ------------------------------------------------------------
static inline uint32_t abs16(int16_t v) {
    return (v < 0) ? (uint32_t)(-(int32_t)v) : (uint32_t)v;
}

static inline uint32_t abs8u(uint8_t v) {
    // 8-bit PCM is unsigned (silence is ~128).
    const int dv = (int)v - 128;
    return (dv < 0) ? (uint32_t)(-dv) : (uint32_t)dv;
}

static inline uint32_t thresholdFor8Bit(uint32_t threshold16) {
    // threshold16 is tuned for 16-bit amplitudes. Map to 8-bit amplitude space (0..127).
    // Dividing by ~64 yields a practical range (~1..3) for typical SoftVoice output.
    uint32_t t = threshold16 / 64u;
    if (t < 1u) t = 1u;
    if (t > 127u) t = 127u;
    return t;
}

static bool isSilentFramePCM16(const uint8_t* frameBytes, int ch, uint32_t threshold16) {
    const int16_t* frame = reinterpret_cast<const int16_t*>(frameBytes);
    for (int c = 0; c < ch; ++c) {
        if (abs16(frame[c]) > threshold16) return false;
    }
    return true;
}

static bool isSilentFramePCM8(const uint8_t* frameBytes, int ch, uint32_t threshold8) {
    for (int c = 0; c < ch; ++c) {
        if (abs8u(frameBytes[c]) > threshold8) return false;
    }
    return true;
}

static size_t computeLeadingTrimBytesLocked(const SV_STATE* s, const StreamItem& it, uint64_t bytesPerSec, int maxTrimMs, int keepMs, uint32_t threshold16) {
    if (!s || !s->formatValid) return 0;
    if (!s->lastFormat.nBlockAlign || !s->lastFormat.nChannels) return 0;
    if (s->lastFormat.wFormatTag != WAVE_FORMAT_PCM) return 0;

    const WORD bits = s->lastFormat.wBitsPerSample;
    if (bits != 8 && bits != 16) return 0;

    const size_t bytesPerSample = (bits == 8) ? 1 : 2;
    const size_t minAlign = (size_t)s->lastFormat.nChannels * bytesPerSample;
    const size_t blockAlign = (size_t)s->lastFormat.nBlockAlign;
    if (blockAlign < minAlign) return 0;

    if (it.offset != 0) return 0; // only before any reads

    const size_t totalFrames = it.data.size() / blockAlign;
    if (totalFrames == 0) return 0;

    const size_t maxFrames = (bytesPerSec && maxTrimMs > 0)
        ? (size_t)((bytesPerSec * (uint64_t)maxTrimMs / 1000ULL) / (uint64_t)blockAlign)
        : totalFrames;

    const size_t keepFrames = (bytesPerSec && keepMs > 0)
        ? (size_t)((bytesPerSec * (uint64_t)keepMs / 1000ULL) / (uint64_t)blockAlign)
        : 0;

    size_t scanFrames = (maxFrames < totalFrames) ? maxFrames : totalFrames;
    if (scanFrames == 0) return 0;

    const int ch = (int)s->lastFormat.nChannels;
    const uint8_t* base = it.data.data();
    const uint32_t threshold8 = (bits == 8) ? thresholdFor8Bit(threshold16) : 0;

    size_t i = 0;
    for (; i < scanFrames; ++i) {
        const uint8_t* frameBytes = base + i * blockAlign;
        const bool silent = (bits == 16)
            ? isSilentFramePCM16(frameBytes, ch, threshold16)
            : isSilentFramePCM8(frameBytes, ch, threshold8);
        if (!silent) break;
    }

    if (i <= keepFrames) return 0;
    const size_t trimFrames = i - keepFrames;
    return trimFrames * blockAlign;
}

static size_t computeTrailingTrimBytesLocked(const SV_STATE* s, const StreamItem& it, uint64_t bytesPerSec, int maxTrimMs, int keepMs, uint32_t threshold16) {
    if (!s || !s->formatValid) return 0;
    if (!s->lastFormat.nBlockAlign || !s->lastFormat.nChannels) return 0;
    if (s->lastFormat.wFormatTag != WAVE_FORMAT_PCM) return 0;

    const WORD bits = s->lastFormat.wBitsPerSample;
    if (bits != 8 && bits != 16) return 0;

    const size_t bytesPerSample = (bits == 8) ? 1 : 2;
    const size_t minAlign = (size_t)s->lastFormat.nChannels * bytesPerSample;
    const size_t blockAlign = (size_t)s->lastFormat.nBlockAlign;
    if (blockAlign < minAlign) return 0;

    const size_t dataSz = it.data.size();
    if (dataSz < blockAlign) return 0;

    // Only trim bytes we have not already handed out.
    const size_t off = it.offset;
    if (off >= dataSz) return 0;

    // Align scan boundaries to whole frames for analysis.
    const size_t scanEnd = (dataSz / blockAlign) * blockAlign;
    if (scanEnd == 0 || off >= scanEnd) return 0;

    // Start scanning after the bytes we've already delivered, rounded up to the next full frame.
    const size_t scanStart = ((off + blockAlign - 1) / blockAlign) * blockAlign;
    if (scanStart >= scanEnd) return 0;

    const size_t totalFrames = scanEnd / blockAlign;
    const size_t startFrame = scanStart / blockAlign;
    const size_t availableFrames = (totalFrames > startFrame) ? (totalFrames - startFrame) : 0;
    if (availableFrames == 0) return 0;

    const size_t maxFrames = (bytesPerSec && maxTrimMs > 0)
        ? (size_t)((bytesPerSec * (uint64_t)maxTrimMs / 1000ULL) / (uint64_t)blockAlign)
        : availableFrames;

    const size_t keepFrames = (bytesPerSec && keepMs > 0)
        ? (size_t)((bytesPerSec * (uint64_t)keepMs / 1000ULL) / (uint64_t)blockAlign)
        : 0;

    size_t scanFrames = (maxFrames < availableFrames) ? maxFrames : availableFrames;
    if (scanFrames == 0) return 0;

    const int ch = (int)s->lastFormat.nChannels;
    const uint8_t* base = it.data.data();
    const uint32_t threshold8 = (bits == 8) ? thresholdFor8Bit(threshold16) : 0;

    // Scan backwards over the unread portion only.
    size_t trailing = 0;
    for (size_t j = 0; j < scanFrames; ++j) {
        const size_t idx = totalFrames - 1 - j;
        if (idx < startFrame) break; // safety
        const uint8_t* frameBytes = base + idx * blockAlign;
        const bool silent = (bits == 16)
            ? isSilentFramePCM16(frameBytes, ch, threshold16)
            : isSilentFramePCM8(frameBytes, ch, threshold8);
        if (!silent) break;
        ++trailing;
    }

    if (trailing <= keepFrames) return 0;
    size_t trimFrames = trailing - keepFrames;
    size_t trimBytes = trimFrames * blockAlign;

    // Never trim past scanStart, and never trim past the unread portion.
    const size_t maxTrimBytes = scanEnd - scanStart;
    if (trimBytes > maxTrimBytes) trimBytes = maxTrimBytes;

    const size_t remaining = dataSz - off;
    if (trimBytes > remaining) trimBytes = remaining;

    return trimBytes;
}

// ------------------------------------------------------------
// Worker loop
// ------------------------------------------------------------
static void workerLoop(SV_STATE* s, int initialVoice) {
    if (!s) return;

    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_ABOVE_NORMAL);

    if (!ensureMsgWindowCreated(s)) {
        s->initOk.store(-1, std::memory_order_relaxed);
        if (s->initEvent) SetEvent(s->initEvent);
        return;
    }

    // Open initial voice.
    if (!openVoiceOnWorker(s, initialVoice)) {
        s->initOk.store(-1, std::memory_order_relaxed);
        if (s->initEvent) SetEvent(s->initEvent);
        if (s->msgWnd) {
            DestroyWindow(s->msgWnd);
            s->msgWnd = nullptr;
        }
        return;
    }

    s->initOk.store(1, std::memory_order_relaxed);
    if (s->initEvent) SetEvent(s->initEvent);

    // Default pacing if format unknown.
    s->bytesPerSec.store(22050, std::memory_order_relaxed);
    computeBufferLimits(s);

    MSG msg;

    auto pumpMessages = [&]() {
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    };

    while (true) {
        pumpMessages();

        // Fetch next command (if any).
        Cmd cmd;
        bool hasCmd = false;
        {
            std::lock_guard<std::mutex> lk(s->cmdMtx);
            if (!s->cmdQ.empty()) {
                cmd = std::move(s->cmdQ.front());
                s->cmdQ.pop_front();
                hasCmd = true;
            } else {
                ResetEvent(s->cmdEvent);
            }
        }

        if (!hasCmd) {
            // Wait for either a command or a message.
            MsgWaitForMultipleObjectsEx(
                1,
                &s->cmdEvent,
                INFINITE,
                QS_ALLINPUT,
                MWMO_INPUTAVAILABLE
            );
            continue;
        }

        if (cmd.type == Cmd::CMD_QUIT) {
            break;
        }

        const uint32_t snap = s->cancelToken.load(std::memory_order_relaxed);
        if (cmd.cancelSnapshot != snap) {
            continue;
        }

        const uint32_t gen = s->genCounter.fetch_add(1, std::memory_order_relaxed);

        ResetEvent(s->stopEvent);
        ResetEvent(s->doneEvent);
        if (s->startEvent) ResetEvent(s->startEvent);

        // gate on
        s->currentGen.store(gen, std::memory_order_relaxed);
        s->activeGen.store(gen, std::memory_order_relaxed);
        s->lastAudioTick.store(0, std::memory_order_relaxed);

        // Reset trim flags (they are per-gen; atomic holds last-done gen)
        // Nothing required here; comparisons in sv_read will handle.

        // clear output
        {
            std::lock_guard<std::mutex> g(s->outMtx);
            clearOutputQueueLocked(s);
        }

        // Ensure correct voice/language.
        bool voiceChanged = false;
        int wantVoice = s->voice.value.load(std::memory_order_relaxed);
        if (wantVoice <= 0) wantVoice = 1;

        if (wantVoice != s->currentVoice) {
            // Prefer SVSetLanguage if available; fallback to reopen.
            if (setLanguageOnWorker(s, wantVoice)) {
                voiceChanged = true;
            } else {
                if (!openVoiceOnWorker(s, wantVoice)) {
                    s->activeGen.store(0, std::memory_order_relaxed);
                    pushMarker(s, SV_ITEM_ERROR, 2003, gen);
                    pushMarker(s, SV_ITEM_DONE, 0, gen);
                    continue;
                }
                voiceChanged = true;
            }
        }

        // Apply personality first (preset), if user selected one.
        const bool personalityApplied = applyPersonalityOnWorker(s, voiceChanged /* reapply on voice switch */);

        // Personalities (variants) act like presets. We *do not* want to stomp their internal params
        // (pitch, wobble, formants, etc) by reapplying the user's sliders every time a variant is chosen.
        //
        // Behavior here aims to match the legacy NVDA SoftVoice driver:
        // - When a non-zero personality is applied, we keep the preset sound and only reapply RATE.
        // - When personality is reset to 0 (back to base), we force-apply all numeric settings so sliders take effect.
        const int persVal = s->personality.value.load(std::memory_order_relaxed);
        const bool persUserSet = (s->personality.userSet.load(std::memory_order_relaxed) != 0);
        const bool persNonZero = persUserSet && (persVal != 0);

        if (personalityApplied && persVal != 0) {
            discardTimbreDirtyOnWorker(s);
        }

        const bool forceNumeric = ((voiceChanged && !persNonZero) || (personalityApplied && persVal == 0));

        // Apply numeric settings (rate, pitch, etc) after personality, but only force when appropriate.
        applyNumericSettingsOnWorker(s, forceNumeric);

        // Legacy behavior: after applying a non-zero personality, reapply the user's rate so speed stays stable.
        if (personalityApplied && persVal != 0) {
            const int rateVal = s->rate.value.load(std::memory_order_relaxed);
            seh_svSet2Int(s->svSetRate, s->svHandle, rateVal);
        }

        // Apply optional style settings only if user explicitly touched them.
        const bool forceStyleIfUserSet = voiceChanged || personalityApplied;

        applyStyleSettingOnWorker(s, s->f0Style, s->svSetF0Style, forceStyleIfUserSet);
        applyStyleSettingOnWorker(s, s->voicingMode, s->svSetVoicingMode, forceStyleIfUserSet);
        applyStyleSettingOnWorker(s, s->gender, s->svSetGender, forceStyleIfUserSet);
        applyStyleSettingOnWorker(s, s->glottalSource, s->svSetGlottalSource, forceStyleIfUserSet);
        applyStyleSettingOnWorker(s, s->speakingMode, s->svSetSpeakingMode, forceStyleIfUserSet);

        // Text conversion
        std::string safe = sanitizeForSoftVoiceCp1252(cmd.text.c_str());
        if (safe.empty()) {
            s->activeGen.store(0, std::memory_order_relaxed);
            pushMarker(s, SV_ITEM_DONE, 0, gen);
            continue;
        }

        // Split long inputs into ~350-char chunks (split on the first space after 350 chars).
        const size_t kChunkChars = 350;
        std::vector<std::string> chunks = splitSoftVoiceTextIntoChunks(safe, kChunkChars);
        if (chunks.empty()) {
            s->activeGen.store(0, std::memory_order_relaxed);
            pushMarker(s, SV_ITEM_DONE, 0, gen);
            continue;
        }

        const auto maxDur = std::chrono::seconds(180);
        HANDLE waits[2] = { s->doneEvent, s->stopEvent };

        bool stopped = false;
        bool ttsError = false;

        for (const std::string& chunk : chunks) {
            if (chunk.empty()) continue;

            if (s->cancelToken.load(std::memory_order_relaxed) != snap) {
                stopped = true;
                break;
            }
            if (WaitForSingleObject(s->stopEvent, 0) == WAIT_OBJECT_0) {
                stopped = true;
                break;
            }

            ResetEvent(s->doneEvent);
            if (s->startEvent) ResetEvent(s->startEvent);

            int rc = seh_svTTS(s->svTTS, s->svHandle, chunk.c_str(), 0, 0, s->msgWnd, 0, 0, 0);
            if (rc != 0) {
                ttsError = true;
                break;
            }

            // Wait for this chunk to finish, while pumping messages.
            const auto t0 = std::chrono::steady_clock::now();
            while (true) {
                DWORD w = MsgWaitForMultipleObjectsEx(
                    2,
                    waits,
                    50,
                    QS_ALLINPUT,
                    MWMO_INPUTAVAILABLE
                );

                if (w == WAIT_OBJECT_0) {
                    // doneEvent
                    break;
                }
                if (w == WAIT_OBJECT_0 + 1) {
                    stopped = true;
                    break;
                }
                if (w == WAIT_OBJECT_0 + 2) {
                    pumpMessages();
                }

                if (s->cancelToken.load(std::memory_order_relaxed) != snap) {
                    stopped = true;
                    break;
                }
                if (std::chrono::steady_clock::now() - t0 > maxDur) {
                    pushMarker(s, SV_ITEM_ERROR, 2002, gen);
                    stopped = true;
                    break;
                }
            }

            if (stopped) break;
        }

        if (ttsError) {
            s->activeGen.store(0, std::memory_order_relaxed);
            pushMarker(s, SV_ITEM_ERROR, 2001, gen);
            pushMarker(s, SV_ITEM_DONE, 0, gen);
            continue;
        }

        if (stopped) {
            // Abort inside worker thread.
            seh_svAbort(s->svAbort, s->svHandle);
            s->activeGen.store(0, std::memory_order_relaxed);
            pushMarker(s, SV_ITEM_DONE, 0, gen);
            continue;
        }

        bool skipGrace = false;
        {
            std::lock_guard<std::mutex> lk(s->cmdMtx);
            skipGrace = !s->cmdQ.empty();
        }

        if (!skipGrace) {
            // Tail-grace: wait until no new audio for ~30ms (max 250ms)
            ULONGLONG graceStart = GetTickCount64();
            while (true) {
                ULONGLONG last = s->lastAudioTick.load(std::memory_order_relaxed);
                ULONGLONG now = GetTickCount64();

                if (last != 0 && (now - last) >= 30) break;
                if ((now - graceStart) >= 250) break;

                if (WaitForSingleObject(s->stopEvent, 5) == WAIT_OBJECT_0) break;
            }
        }

        // gate off before DONE
        s->activeGen.store(0, std::memory_order_relaxed);
        pushMarker(s, SV_ITEM_DONE, 0, gen);
    }

    // Cleanup on worker thread.
    if (s->svHandle != 0) {
        seh_svAbort(s->svAbort, s->svHandle);
        seh_svCloseSpeech(s->svCloseSpeech, s->svHandle);
        s->svHandle = 0;
    }
    if (s->msgWnd) {
        DestroyWindow(s->msgWnd);
        s->msgWnd = nullptr;
    }
}

// ------------------------------------------------------------
// Exports
// ------------------------------------------------------------
extern "C" SV_API SV_STATE* __cdecl sv_initW(const wchar_t* baseDllPath, int initialVoice) {
    if (!baseDllPath) return nullptr;

    // Singleton + refcount: return existing instance if already initialized.
    // This avoids "wrapper init failed" when NVDA briefly loads a new synth instance
    // before the old one is terminated.
    std::lock_guard<std::mutex> glk(g_globalMtx);
    if (g_state) {
        // Reuse the global instance (NVDA may create multiple driver instances across synth switches).
        // We intentionally do *not* refuse re-init based on path mismatches; doing so can prevent reloading.
        if (g_state->baseDllPath.empty()) {
            g_state->baseDllPath = baseDllPath;
        }
        g_refCount.fetch_add(1, std::memory_order_relaxed);
        return g_state;
    }

    HMODULE base = LoadLibraryW(baseDllPath);
    if (!base) return nullptr;

    // Try to load language DLLs from the same folder (non-fatal if missing).
    std::wstring dir;
    {
        std::wstring p(baseDllPath);
        size_t pos = p.find_last_of(L"\\/");
        if (pos != std::wstring::npos) dir = p.substr(0, pos);
    }

    HMODULE eng = nullptr;
    HMODULE span = nullptr;
    if (!dir.empty()) {
        std::wstring pEng = dir + L"\\tieng32.dll";
        std::wstring pSpan = dir + L"\\tispan32.dll";
        eng = LoadLibraryW(pEng.c_str());
        span = LoadLibraryW(pSpan.c_str());
    }

    auto* s = new SV_STATE();
    s->baseDllPath = baseDllPath;
    s->baseModule = base;
    s->engModule = eng;
    s->spanModule = span;

    s->svSyncMsg = RegisterWindowMessageW(L"SVSyncMessages");
    s->activeSyncMsg = 0;

    s->svOpenSpeech = (SVOpenSpeechFunc)getProcMaybeDecorated(base, "SVOpenSpeech", "_SVOpenSpeech@20");
    s->svCloseSpeech = (SVCloseSpeechFunc)getProcMaybeDecorated(base, "SVCloseSpeech", "_SVCloseSpeech@4");
    s->svAbort = (SVAbortFunc)getProcMaybeDecorated(base, "SVAbort", "_SVAbort@4");
    s->svTTS = (SVTTSFunc)getProcMaybeDecorated(base, "SVTTS", "_SVTTS@32");

    // optional
    s->svSetLanguage = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetLanguage", "_SVSetLanguage@8");

    s->svSetRate = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetRate", "_SVSetRate@8");
    s->svSetPitch = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetPitch", "_SVSetPitch@8");
    s->svSetF0Range = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetF0Range", "_SVSetF0Range@8");
    s->svSetF0Perturb = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetF0Perturb", "_SVSetF0Perturb@8");
    s->svSetVowelFactor = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetVowelFactor", "_SVSetVowelFactor@8");

    s->svSetAVBias = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetAVBias", "_SVSetAVBias@8");
    s->svSetAFBias = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetAFBias", "_SVSetAFBias@8");
    s->svSetAHBias = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetAHBias", "_SVSetAHBias@8");

    s->svSetPersonality = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetPersonality", "_SVSetPersonality@8");
    s->svSetF0Style = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetF0Style", "_SVSetF0Style@8");
    s->svSetVoicingMode = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetVoicingMode", "_SVSetVoicingMode@8");
    s->svSetGender = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetGender", "_SVSetGender@8");
    s->svSetGlottalSource = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetGlottalSource", "_SVSetGlottalSource@8");
    s->svSetSpeakingMode = (SVSet2IntFunc)getProcMaybeDecorated(base, "SVSetSpeakingMode", "_SVSetSpeakingMode@8");

    s->startEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->doneEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->cmdEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->initEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);

    if (!s->startEvent || !s->doneEvent || !s->stopEvent || !s->cmdEvent || !s->initEvent
        || !s->svOpenSpeech || !s->svCloseSpeech || !s->svAbort || !s->svTTS) {

        if (s->startEvent) CloseHandle(s->startEvent);
        if (s->doneEvent) CloseHandle(s->doneEvent);
        if (s->stopEvent) CloseHandle(s->stopEvent);
        if (s->cmdEvent) CloseHandle(s->cmdEvent);
        if (s->initEvent) CloseHandle(s->initEvent);

        if (span) FreeLibrary(span);
        if (eng) FreeLibrary(eng);
        FreeLibrary(base);
        delete s;
        return nullptr;
    }

    // Defaults: these roughly match the legacy driver mapping.
    // Numeric defaults are considered user-set so we apply them (dirty=1) at least once.
    s->voice.value.store((initialVoice > 0) ? initialVoice : 1, std::memory_order_relaxed);
    s->voice.userSet.store(1, std::memory_order_relaxed);
    s->voice.dirty.store(0, std::memory_order_relaxed);

    s->rate.value.store(260, std::memory_order_relaxed);
    s->rate.userSet.store(1, std::memory_order_relaxed);
    s->rate.dirty.store(1, std::memory_order_relaxed);

    s->pitch.value.store(89, std::memory_order_relaxed);
    s->pitch.userSet.store(1, std::memory_order_relaxed);
    s->pitch.dirty.store(1, std::memory_order_relaxed);

    s->f0Range.value.store(125, std::memory_order_relaxed);
    s->f0Range.userSet.store(1, std::memory_order_relaxed);
    s->f0Range.dirty.store(1, std::memory_order_relaxed);

    s->f0Perturb.value.store(0, std::memory_order_relaxed);
    s->f0Perturb.userSet.store(1, std::memory_order_relaxed);
    s->f0Perturb.dirty.store(1, std::memory_order_relaxed);

    s->vowelFactor.value.store(100, std::memory_order_relaxed);
    s->vowelFactor.userSet.store(1, std::memory_order_relaxed);
    s->vowelFactor.dirty.store(1, std::memory_order_relaxed);

    s->avBias.value.store(0, std::memory_order_relaxed);
    s->avBias.userSet.store(1, std::memory_order_relaxed);
    s->avBias.dirty.store(1, std::memory_order_relaxed);

    s->afBias.value.store(0, std::memory_order_relaxed);
    s->afBias.userSet.store(1, std::memory_order_relaxed);
    s->afBias.dirty.store(1, std::memory_order_relaxed);

    s->ahBias.value.store(0, std::memory_order_relaxed);
    s->ahBias.userSet.store(1, std::memory_order_relaxed);
    s->ahBias.dirty.store(1, std::memory_order_relaxed);

    // Personality + style params are NOT user-set by default (so voices like Robot/Martian can use presets).
    s->personality.value.store(0, std::memory_order_relaxed);
    s->personality.userSet.store(0, std::memory_order_relaxed);
    s->personality.dirty.store(0, std::memory_order_relaxed);

    s->f0Style.value.store(0, std::memory_order_relaxed);
    s->f0Style.userSet.store(0, std::memory_order_relaxed);
    s->f0Style.dirty.store(0, std::memory_order_relaxed);

    s->voicingMode.value.store(0, std::memory_order_relaxed);
    s->voicingMode.userSet.store(0, std::memory_order_relaxed);
    s->voicingMode.dirty.store(0, std::memory_order_relaxed);

    s->gender.value.store(0, std::memory_order_relaxed);
    s->gender.userSet.store(0, std::memory_order_relaxed);
    s->gender.dirty.store(0, std::memory_order_relaxed);

    s->glottalSource.value.store(0, std::memory_order_relaxed);
    s->glottalSource.userSet.store(0, std::memory_order_relaxed);
    s->glottalSource.dirty.store(0, std::memory_order_relaxed);

    s->speakingMode.value.store(0, std::memory_order_relaxed);
    s->speakingMode.userSet.store(0, std::memory_order_relaxed);
    s->speakingMode.dirty.store(0, std::memory_order_relaxed);

    // Lead defaults
    s->maxLeadMs.store(2000, std::memory_order_relaxed);
    s->autoLead.store(1, std::memory_order_relaxed);

    s->baseDllPath = baseDllPath;

    g_state = s;
    g_refCount.store(1, std::memory_order_relaxed);
    ensureHooksInstalled();

    int v = (initialVoice > 0) ? initialVoice : 1;
    s->worker = std::thread(workerLoop, s, v);

    // Wait for the worker to finish initial setup (message window + SVOpenSpeech).
    DWORD w = WaitForSingleObject(s->initEvent, 5000);
    const int ok = s->initOk.load(std::memory_order_relaxed);
    CloseHandle(s->initEvent);
    s->initEvent = nullptr;

    if (w != WAIT_OBJECT_0 || ok != 1) {
        // Best effort cleanup.
        s->cancelToken.fetch_add(1, std::memory_order_relaxed);
        SetEvent(s->stopEvent);
        SetEvent(s->doneEvent);

        {
            std::lock_guard<std::mutex> lk(s->cmdMtx);
            s->cmdQ.clear();
            Cmd q;
            q.type = Cmd::CMD_QUIT;
            q.cancelSnapshot = s->cancelToken.load(std::memory_order_relaxed);
            s->cmdQ.push_back(std::move(q));
        }
        SetEvent(s->cmdEvent);

        if (s->worker.joinable()) s->worker.join();

        if (s->startEvent) CloseHandle(s->startEvent);
        if (s->doneEvent) CloseHandle(s->doneEvent);
        if (s->stopEvent) CloseHandle(s->stopEvent);
        if (s->cmdEvent) CloseHandle(s->cmdEvent);

        { auto fl = [](HMODULE& m) { if (!m) return; while (FreeLibrary(m)) {} m = nullptr; };
          fl(s->spanModule); fl(s->engModule); fl(s->baseModule); }

        if (g_state == s) g_state = nullptr;
        delete s;
        return nullptr;
    }
    return s;
}

extern "C" SV_API void __cdecl sv_free(SV_STATE* s) {
    if (!s) return;

    // Refcounted singleton. Only tear down when the last caller releases.
    // Keep g_state valid during teardown so the hooks continue to swallow SoftVoice's
    // waveOut* calls until the engine is fully stopped.
    std::lock_guard<std::mutex> glk(g_globalMtx);
    if (s != g_state) return;
    long rc = g_refCount.load(std::memory_order_relaxed);
    if (rc > 1) {
        g_refCount.store(rc - 1, std::memory_order_relaxed);
        return;
    }
    g_refCount.store(0, std::memory_order_relaxed);

    // cancel + wake everything
    s->cancelToken.fetch_add(1, std::memory_order_relaxed);
    SetEvent(s->stopEvent);
    SetEvent(s->doneEvent);
    if (s->startEvent) SetEvent(s->startEvent);

    s->activeGen.store(0, std::memory_order_relaxed);
    s->currentGen.store(0, std::memory_order_relaxed);

    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        s->cmdQ.clear();
        Cmd q;
        q.type = Cmd::CMD_QUIT;
        q.cancelSnapshot = s->cancelToken.load(std::memory_order_relaxed);
        s->cmdQ.push_back(std::move(q));
    }
    SetEvent(s->cmdEvent);

    if (s->worker.joinable()) s->worker.join();

    {
        std::lock_guard<std::mutex> g(s->outMtx);
        clearOutputQueueLocked(s);
    }

    if (s->startEvent) CloseHandle(s->startEvent);
    if (s->doneEvent) CloseHandle(s->doneEvent);
    if (s->stopEvent) CloseHandle(s->stopEvent);
    if (s->cmdEvent) CloseHandle(s->cmdEvent);

    // Force-unload the SoftVoice engine DLLs by draining any extra
    // LoadLibrary references.  tibase32.dll is a late-90s engine that
    // keeps internal global state which is NOT reset by SVCloseSpeech.
    // A single FreeLibrary may leave the DLL mapped (refcount > 0) if
    // the engine or its dependencies called LoadLibrary internally.
    // Looping until the module is truly gone ensures a future sv_initW
    // gets a pristine DLL_PROCESS_ATTACH.
    auto forceUnload = [](HMODULE& mod) {
        if (!mod) return;
        // FreeLibrary returns FALSE once the module is already gone.
        while (FreeLibrary(mod)) {}
        mod = nullptr;
    };
    forceUnload(s->spanModule);
    forceUnload(s->engModule);
    forceUnload(s->baseModule);

    if (g_state == s) g_state = nullptr;
    delete s;
}

extern "C" SV_API void __cdecl sv_stop(SV_STATE* s) {
    if (!s) return;

    s->cancelToken.fetch_add(1, std::memory_order_relaxed);

    // gate off and clear queue
    s->activeGen.store(0, std::memory_order_relaxed);
    s->currentGen.store(0, std::memory_order_relaxed);

    {
        std::lock_guard<std::mutex> g(s->outMtx);
        clearOutputQueueLocked(s);
    }

    // Clear pending commands.
    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        s->cmdQ.clear();
    }

    // wake worker + hook waits
    SetEvent(s->stopEvent);
    SetEvent(s->doneEvent);
    if (s->startEvent) SetEvent(s->startEvent);
    SetEvent(s->cmdEvent);
}

extern "C" SV_API int __cdecl sv_startSpeakW(SV_STATE* s, const wchar_t* text) {
    if (!s || !text) return 1;

    Cmd cmd;
    cmd.type = Cmd::CMD_SPEAK;
    cmd.cancelSnapshot = s->cancelToken.load(std::memory_order_relaxed);
    cmd.text = text;

    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        s->cmdQ.push_back(std::move(cmd));
    }
    SetEvent(s->cmdEvent);
    return 0;
}

extern "C" SV_API int __cdecl sv_read(SV_STATE* s, int* outType, int* outValue, uint8_t* outAudio, int outCap) {
    if (outType) *outType = SV_ITEM_NONE;
    if (outValue) *outValue = 0;
    if (!s || !outAudio || outCap < 0) return 0;

    std::lock_guard<std::mutex> g(s->outMtx);

    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0) {
        clearOutputQueueLocked(s);
        return 0;
    }

    // drop stale items
    while (!s->outQ.empty() && s->outQ.front().gen != curGen) {
        StreamItem& it = s->outQ.front();
        if (it.type == SV_ITEM_AUDIO) {
            size_t remaining = (it.data.size() > it.offset) ? (it.data.size() - it.offset) : 0;
            if (s->queuedAudioBytes >= remaining) s->queuedAudioBytes -= remaining;
            else s->queuedAudioBytes = 0;
        }
        s->outQ.pop_front();
    }

    if (s->outQ.empty()) return 0;

    // Optional silence trimming (only when enabled).
    // We do this here because:
    // - It's safe (no impact on SoftVoice), just affects what we hand to NVWave.
    // - It can reduce chunk-boundary pauses that SoftVoice sometimes emits as silence.
    if (s->trimSilence.load(std::memory_order_relaxed) != 0 && s->formatValid) {
        const uint64_t bps = s->bytesPerSec.load(std::memory_order_relaxed);
        // Conservative parameters, scaled by pauseFactor.
        // pauseFactor=0 => very light trim
        // pauseFactor=100 => heavier trim (still leaves a small safety tail)
        int pf = s->pauseFactor.load(std::memory_order_relaxed);
        if (pf < 0) pf = 0;
        if (pf > 100) pf = 100;

        const int maxTrimLeadMs = 200 + (pf * 12);   // 200..1400
        const int keepLeadMs = 8;                  // keep a little audio to avoid clipping
        const int maxTrimTailMs = 250 + (pf * 12);  // 250..1450
        const int keepTailMs = 10;                 // keep a little tail for consonants
        const uint32_t threshold = 48 + (uint32_t)(pf * 2); // abs(sample) <= threshold treated as silence

        // Trim leading silence once per gen.
        if (s->leadTrimDoneGen.load(std::memory_order_relaxed) != curGen) {
            // Find first audio item (normally front).
            for (auto it = s->outQ.begin(); it != s->outQ.end(); ++it) {
                if (it->type != SV_ITEM_AUDIO) continue;
                size_t trim = computeLeadingTrimBytesLocked(s, *it, bps, maxTrimLeadMs, keepLeadMs, threshold);
                if (trim > 0) {
                    if (trim > it->data.size()) trim = it->data.size();
                    it->offset += trim;
                    if (s->queuedAudioBytes >= trim) s->queuedAudioBytes -= trim;
                    else s->queuedAudioBytes = 0;
                }
                break;
            }
            s->leadTrimDoneGen.store(curGen, std::memory_order_relaxed);

            // Drop any audio items that became empty.
            while (!s->outQ.empty() && s->outQ.front().type == SV_ITEM_AUDIO && s->outQ.front().offset >= s->outQ.front().data.size()) {
                s->outQ.pop_front();
            }
            if (s->outQ.empty()) return 0;
        }

        // Trim trailing silence once per gen, but only once DONE is queued.
        if (s->tailTrimDoneGen.load(std::memory_order_relaxed) != curGen) {
            // Only attempt if we can see a DONE marker in the queue (usually at the end).
            bool hasDone = false;
            for (auto rit = s->outQ.rbegin(); rit != s->outQ.rend(); ++rit) {
                if (rit->type == SV_ITEM_DONE) { hasDone = true; break; }
                if (rit->type == SV_ITEM_ERROR) continue;
                // if there are no markers at end we still might be in-progress; keep scanning
            }
            if (hasDone) {
                // Find the last audio item (unconsumed) and trim its tail.
                for (auto rit = s->outQ.rbegin(); rit != s->outQ.rend(); ++rit) {
                    if (rit->type != SV_ITEM_AUDIO) continue;

                    const size_t oldSz = rit->data.size();
                    const size_t oldOff = rit->offset;

                    size_t trim = computeTrailingTrimBytesLocked(s, *rit, bps, maxTrimTailMs, keepTailMs, threshold);
                    if (trim > 0 && oldSz > oldOff) {
                        const size_t remaining = oldSz - oldOff;
                        if (trim > remaining) trim = remaining;

                        if (trim > 0) {
                            const size_t newSz = oldSz - trim;
                            if (newSz >= oldOff) {
                                rit->data.resize(newSz);
                                // Update queued bytes: we removed trim bytes that would have been delivered.
                                if (s->queuedAudioBytes >= trim) s->queuedAudioBytes -= trim;
                                else s->queuedAudioBytes = 0;
                            }
                        }
                    }
                    break;
                }
                s->tailTrimDoneGen.store(curGen, std::memory_order_relaxed);
            }
        }
    }

    if (s->outQ.empty()) return 0;

    StreamItem& front = s->outQ.front();
    if (outType) *outType = front.type;
    if (outValue) *outValue = front.value;

    if (front.type == SV_ITEM_AUDIO) {
        size_t remainingSz = (front.data.size() > front.offset) ? (front.data.size() - front.offset) : 0;
        int remaining = (remainingSz > (size_t)INT_MAX) ? INT_MAX : (int)remainingSz;
        int n = remaining;
        if (n > outCap) n = outCap;

        if (n > 0) {
            std::memcpy(outAudio, front.data.data() + front.offset, (size_t)n);
            front.offset += (size_t)n;
            if (s->queuedAudioBytes >= (size_t)n) s->queuedAudioBytes -= (size_t)n;
            else s->queuedAudioBytes = 0;
        }

        if (front.offset >= front.data.size()) {
            s->outQ.pop_front();
        }
        return n;
    }

    // DONE/ERROR
    s->outQ.pop_front();
    return 0;
}

// ------------------------------------------------------------
// Settings API: store desired values; worker applies them.
// ------------------------------------------------------------
static void setSetting(SettingInt& st, int v, bool isUserSet) {
    st.value.store(v, std::memory_order_relaxed);
    st.dirty.store(1, std::memory_order_relaxed);
    if (isUserSet) st.userSet.store(1, std::memory_order_relaxed);
}

extern "C" SV_API int __cdecl sv_getRate(SV_STATE* s) { return s ? s->rate.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setRate(SV_STATE* s, int v) { if (s) setSetting(s->rate, v, true); }

extern "C" SV_API int __cdecl sv_getPitch(SV_STATE* s) { return s ? s->pitch.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setPitch(SV_STATE* s, int v) { if (s) setSetting(s->pitch, v, true); }

extern "C" SV_API int __cdecl sv_getF0Range(SV_STATE* s) { return s ? s->f0Range.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setF0Range(SV_STATE* s, int v) { if (s) setSetting(s->f0Range, v, true); }

extern "C" SV_API int __cdecl sv_getF0Perturb(SV_STATE* s) { return s ? s->f0Perturb.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setF0Perturb(SV_STATE* s, int v) { if (s) setSetting(s->f0Perturb, v, true); }

extern "C" SV_API int __cdecl sv_getVowelFactor(SV_STATE* s) { return s ? s->vowelFactor.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setVowelFactor(SV_STATE* s, int v) { if (s) setSetting(s->vowelFactor, v, true); }

extern "C" SV_API int __cdecl sv_getAVBias(SV_STATE* s) { return s ? s->avBias.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setAVBias(SV_STATE* s, int v) { if (s) setSetting(s->avBias, v, true); }

extern "C" SV_API int __cdecl sv_getAFBias(SV_STATE* s) { return s ? s->afBias.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setAFBias(SV_STATE* s, int v) { if (s) setSetting(s->afBias, v, true); }

extern "C" SV_API int __cdecl sv_getAHBias(SV_STATE* s) { return s ? s->ahBias.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setAHBias(SV_STATE* s, int v) { if (s) setSetting(s->ahBias, v, true); }

extern "C" SV_API int __cdecl sv_getPersonality(SV_STATE* s) { return s ? s->personality.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setPersonality(SV_STATE* s, int v) { if (s) setSetting(s->personality, v, true); }

extern "C" SV_API int __cdecl sv_getF0Style(SV_STATE* s) { return s ? s->f0Style.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setF0Style(SV_STATE* s, int v) { if (s) setSetting(s->f0Style, v, true); }

extern "C" SV_API int __cdecl sv_getVoicingMode(SV_STATE* s) { return s ? s->voicingMode.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setVoicingMode(SV_STATE* s, int v) { if (s) setSetting(s->voicingMode, v, true); }

extern "C" SV_API int __cdecl sv_getGender(SV_STATE* s) { return s ? s->gender.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setGender(SV_STATE* s, int v) { if (s) setSetting(s->gender, v, true); }

extern "C" SV_API int __cdecl sv_getGlottalSource(SV_STATE* s) { return s ? s->glottalSource.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setGlottalSource(SV_STATE* s, int v) { if (s) setSetting(s->glottalSource, v, true); }

extern "C" SV_API int __cdecl sv_getSpeakingMode(SV_STATE* s) { return s ? s->speakingMode.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setSpeakingMode(SV_STATE* s, int v) {
    if (!s) return;
    setSetting(s->speakingMode, v, true);

    // Auto-tune lead: word-at-a-time/spelled are easier to keep correct if we don't synth far ahead.
    if (s->autoLead.load(std::memory_order_relaxed) != 0) {
        int lead = 2000;
        // Historically we forced lead=0 for word/spell modes, but that can exaggerate
        // perceived choppiness. Keep a small lead while still being fairly "locked".
        if (v == 1 || v == 2) lead = 250;
        s->maxLeadMs.store(lead, std::memory_order_relaxed);
    }
}

extern "C" SV_API int __cdecl sv_getVoice(SV_STATE* s) { return s ? s->voice.value.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setVoice(SV_STATE* s, int v) { if (s) setSetting(s->voice, v, true); }

// Optional knobs
extern "C" SV_API int __cdecl sv_getMaxLeadMs(SV_STATE* s) { return s ? s->maxLeadMs.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setMaxLeadMs(SV_STATE* s, int maxLeadMs) {
    if (!s) return;
    if (maxLeadMs < 0) maxLeadMs = 0;
    if (maxLeadMs > 15000) maxLeadMs = 15000;
    s->autoLead.store(0, std::memory_order_relaxed);
    s->maxLeadMs.store(maxLeadMs, std::memory_order_relaxed);
}

extern "C" SV_API int __cdecl sv_getTrimSilence(SV_STATE* s) { return s ? s->trimSilence.load(std::memory_order_relaxed) : 0; }
extern "C" SV_API void __cdecl sv_setTrimSilence(SV_STATE* s, int enable) {
    if (!s) return;
    s->trimSilence.store(enable ? 1 : 0, std::memory_order_relaxed);
}

extern "C" SV_API int __cdecl sv_getPauseFactor(SV_STATE* s) {
    return s ? s->pauseFactor.load(std::memory_order_relaxed) : 0;
}

extern "C" SV_API void __cdecl sv_setPauseFactor(SV_STATE* s, int factor) {
    if (!s) return;
    if (factor < 0) factor = 0;
    if (factor > 100) factor = 100;
    s->pauseFactor.store(factor, std::memory_order_relaxed);
}

extern "C" SV_API int __cdecl sv_getFormat(SV_STATE* s, int* sampleRate, int* channels, int* bitsPerSample) {
    if (!s || !s->formatValid) return 0;
    if (sampleRate) *sampleRate = (int)s->lastFormat.nSamplesPerSec;
    if (channels) *channels = (int)s->lastFormat.nChannels;
    if (bitsPerSample) *bitsPerSample = (int)s->lastFormat.wBitsPerSample;
    return 1;
}

BOOL APIENTRY DllMain(HMODULE, DWORD, LPVOID) {
    return TRUE;
}
