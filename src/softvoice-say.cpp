// softvoice-say.cpp
//
// Small standalone “Speak Window” for the SoftVoice NVDA wrapper.
// - Enter text (or load a .txt file)
// - Adjust the same key params exposed in the NVDA synth driver
// - Speak out loud, or save to WAV (11025 Hz, PCM, mono)
//
// Build notes (MSVC):
// - Build x86 (SoftVoice is 32-bit).
// - Put these next to the EXE:
//     softvoice_wrapper.dll
//     tibase32.dll (and any related SoftVoice language DLLs)
//
// This file expects a dialog resource (softvoice-say.rc) and resource.h.

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <mmsystem.h>

#include <atomic>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwctype>
#include <fstream>
#include <mutex>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "resource.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "winmm.lib")

// Wrapper stream item types (must match softvoice_wrapper.dll)
static constexpr int SV_ITEM_NONE  = 0;
static constexpr int SV_ITEM_AUDIO = 1;
static constexpr int SV_ITEM_DONE  = 2;
static constexpr int SV_ITEM_ERROR = 3;

static constexpr size_t MAX_SOFTVOICE_CHUNK = 200;
static constexpr int TARGET_WAV_RATE = 11025;
static constexpr int TARGET_WAV_CHANNELS = 1;
static constexpr int TARGET_WAV_BITS = 16;

// -----------------------------------------------------------------------------
// Small helpers
// -----------------------------------------------------------------------------

static std::wstring GetExeDir() {
    wchar_t buf[MAX_PATH] = {};
    DWORD n = GetModuleFileNameW(nullptr, buf, (DWORD)_countof(buf));
    if (n == 0 || n >= (DWORD)_countof(buf)) return L"";
    std::wstring p(buf);
    size_t slash = p.find_last_of(L"\\/");
    if (slash == std::wstring::npos) return L"";
    return p.substr(0, slash);
}

static bool FileExists(const std::wstring& path) {
    DWORD attrs = GetFileAttributesW(path.c_str());
    return (attrs != INVALID_FILE_ATTRIBUTES) && ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0);
}

static std::wstring BaseName(const std::wstring& path) {
    size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos) return path;
    return path.substr(slash + 1);
}

static bool IsTibase32Path(const std::wstring& path) {
    const std::wstring bn = BaseName(path);
    // Case-insensitive compare (Win32)
    return (lstrcmpiW(bn.c_str(), L"tibase32.dll") == 0);
}


static int ClampInt(int v, int lo, int hi) {
    if (v < lo) return lo;
    if (v > hi) return hi;
    return v;
}

static int PercentToParam(int percent, int minVal, int maxVal) {
    percent = ClampInt(percent, 0, 100);
    const double ratio = (double)percent / 100.0;
    const double v = (double)minVal + ((double)maxVal - (double)minVal) * ratio;
    return (int)std::llround(v);
}

static std::wstring Trim(const std::wstring& s) {
    size_t a = 0;
    while (a < s.size() && iswspace(s[a])) a++;
    size_t b = s.size();
    while (b > a && iswspace(s[b - 1])) b--;
    return s.substr(a, b - a);
}

static std::wstring CollapseWhitespaceToSpaces(const std::wstring& s) {
    std::wstring out;
    out.reserve(s.size());
    bool inWs = false;
    for (wchar_t ch : s) {
        if (iswspace(ch)) {
            if (!inWs) {
                out.push_back(L' ');
                inWs = true;
            }
        } else {
            out.push_back(ch);
            inWs = false;
        }
    }
    return Trim(out);
}

static std::wstring SanitizeTextForUi(const std::wstring& s) {
    // Light clean-up, keeping it close to the NVDA driver behavior.
    if (s.empty()) return L"";
    std::wstring t;
    t.reserve(s.size());

    auto mapChar = [](wchar_t c) -> wchar_t {
        switch (c) {
        case 0x2018: case 0x2019: return L'\'';  // ‘ ’
        case 0x201C: case 0x201D: return L'"';   // “ ”
        case 0x2013: case 0x2014: return L'-';   // – —
        case 0x2026: return L'.';                // … (we’ll collapse whitespace later)
        case 0x00A0: return L' ';                // NBSP
        default: return c;
        }
    };

    for (size_t i = 0; i < s.size(); i++) {
        wchar_t c = mapChar(s[i]);

        // Strip a few annoying invisible chars (same family as NVDA driver).
        if (c == 0xFEFF || c == 0x00AD || c == 0x200B || c == 0x200C || c == 0x200D || c == 0x200E || c == 0x200F) {
            continue;
        }

        // Replace C0 controls (except tab/newline) with space.
        if (c < 0x20 && c != L'\t' && c != L'\n' && c != L'\r') {
            c = L' ';
        }

        // Replace surrogate halves with space (SoftVoice is BMP-friendly anyway).
        if (c >= 0xD800 && c <= 0xDFFF) {
            c = L' ';
        }

        // Treat newlines like spaces for speech.
        if (c == L'\r' || c == L'\n' || c == L'\t') c = L' ';
        t.push_back(c);
    }
    return CollapseWhitespaceToSpaces(t);
}

static bool IsAsciiAlphaNum(wchar_t c) {
    return (c >= L'0' && c <= L'9') || (c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z');
}

static std::wstring SpellAlphaNumRuns(const std::wstring& s) {
    // Turn “Hello123” into “H e l l o 1 2 3” (NVDA “Spelled” mode style).
    std::wstring out;
    out.reserve(s.size() * 2);
    size_t i = 0;
    while (i < s.size()) {
        if (!IsAsciiAlphaNum(s[i])) {
            out.push_back(s[i]);
            i++;
            continue;
        }
        size_t j = i;
        while (j < s.size() && IsAsciiAlphaNum(s[j])) j++;
        for (size_t k = i; k < j; k++) {
            out.push_back(s[k]);
            if (k + 1 < j) out.push_back(L' ');
        }
        i = j;
    }
    return CollapseWhitespaceToSpaces(out);
}

static std::vector<std::wstring> SplitForSoftVoice(const std::wstring& text, int speakingMode /*0..2*/) {
    std::vector<std::wstring> out;
    std::wstring t = SanitizeTextForUi(text);
    if (t.empty()) return out;

    if (speakingMode == 2) {
        t = SpellAlphaNumRuns(t);
    }

    if (speakingMode == 1) {
        // Word-at-a-time.
        size_t i = 0;
        while (i < t.size()) {
            while (i < t.size() && t[i] == L' ') i++;
            if (i >= t.size()) break;
            size_t j = i;
            while (j < t.size() && t[j] != L' ') j++;
            std::wstring w = t.substr(i, j - i);
            if (!w.empty()) out.push_back(std::move(w));
            i = j;
        }
        return out;
    }

    // Normal modes: chunk into small pieces at word boundaries.
    std::wstring current;
    current.reserve(MAX_SOFTVOICE_CHUNK);

    size_t i = 0;
    while (i < t.size()) {
        while (i < t.size() && t[i] == L' ') i++;
        if (i >= t.size()) break;
        size_t j = i;
        while (j < t.size() && t[j] != L' ') j++;
        std::wstring word = t.substr(i, j - i);

        if (word.size() > MAX_SOFTVOICE_CHUNK) {
            if (!current.empty()) {
                out.push_back(std::move(current));
                current.clear();
            }
            for (size_t pos = 0; pos < word.size(); pos += MAX_SOFTVOICE_CHUNK) {
                out.push_back(word.substr(pos, MAX_SOFTVOICE_CHUNK));
            }
            i = j;
            continue;
        }

        if (current.empty()) {
            current = std::move(word);
        } else if (current.size() + 1 + word.size() <= MAX_SOFTVOICE_CHUNK) {
            current.push_back(L' ');
            current.append(word);
        } else {
            out.push_back(std::move(current));
            current = std::move(word);
        }
        i = j;
    }
    if (!current.empty()) {
        out.push_back(std::move(current));
    }
    return out;
}

static std::wstring BrowseForFile(HWND owner, bool saveDialog, const wchar_t* title, const wchar_t* filter, const wchar_t* defExt) {
    wchar_t pathBuf[MAX_PATH] = {};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFile = pathBuf;
    ofn.nMaxFile = (DWORD)_countof(pathBuf);
    ofn.lpstrFilter = filter;
    ofn.lpstrTitle = title;
    ofn.Flags = OFN_EXPLORER | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST;
    if (!saveDialog) ofn.Flags |= OFN_FILEMUSTEXIST;
    if (defExt) ofn.lpstrDefExt = defExt;
    BOOL ok = saveDialog ? GetSaveFileNameW(&ofn) : GetOpenFileNameW(&ofn);
    if (!ok) return L"";
    return std::wstring(pathBuf);
}

static bool ReadWholeFileBytes(const std::wstring& path, std::vector<uint8_t>& out) {
    out.clear();
    std::ifstream f(path, std::ios::binary);
    if (!f) return false;
    f.seekg(0, std::ios::end);
    std::streamoff sz = f.tellg();
    if (sz < 0) return false;
    f.seekg(0, std::ios::beg);
    out.resize((size_t)sz);
    if (sz > 0) f.read(reinterpret_cast<char*>(out.data()), sz);
    return true;
}

static std::wstring BytesToWideBestEffort(const std::vector<uint8_t>& bytes) {
    if (bytes.empty()) return L"";

    // UTF-16 LE BOM
    if (bytes.size() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
        const size_t wcharCount = (bytes.size() - 2) / 2;
        std::wstring w;
        w.resize(wcharCount);
        std::memcpy(w.data(), bytes.data() + 2, wcharCount * 2);
        return w;
    }
    // UTF-16 BE BOM
    if (bytes.size() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) {
        const size_t wcharCount = (bytes.size() - 2) / 2;
        std::wstring w;
        w.resize(wcharCount);
        for (size_t i = 0; i < wcharCount; i++) {
            const uint8_t hi = bytes[2 + i * 2];
            const uint8_t lo = bytes[2 + i * 2 + 1];
            w[i] = (wchar_t)((hi << 8) | lo);
        }
        return w;
    }

    // UTF-8 BOM
    size_t start = 0;
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) start = 3;

    auto tryDecode = [&](UINT cp) -> std::wstring {
        const char* src = reinterpret_cast<const char*>(bytes.data() + start);
        int srcLen = (int)(bytes.size() - start);
        int need = MultiByteToWideChar(cp, MB_ERR_INVALID_CHARS, src, srcLen, nullptr, 0);
        if (need <= 0) return L"";
        std::wstring w;
        w.resize((size_t)need);
        MultiByteToWideChar(cp, MB_ERR_INVALID_CHARS, src, srcLen, w.data(), need);
        return w;
    };

    std::wstring w = tryDecode(CP_UTF8);
    if (!w.empty()) return w;

    // Fallback to ANSI codepage.
    const char* src = reinterpret_cast<const char*>(bytes.data() + start);
    int srcLen = (int)(bytes.size() - start);
    int need = MultiByteToWideChar(CP_ACP, 0, src, srcLen, nullptr, 0);
    if (need <= 0) return L"";
    w.resize((size_t)need);
    MultiByteToWideChar(CP_ACP, 0, src, srcLen, w.data(), need);
    return w;
}

// -----------------------------------------------------------------------------
// WAV writing (PCM)
// -----------------------------------------------------------------------------

#pragma pack(push, 1)
struct WavHeader {
    char riff[4];
    uint32_t riffSize;
    char wave[4];
    char fmt_[4];
    uint32_t fmtSize;
    uint16_t audioFormat;
    uint16_t numChannels;
    uint32_t sampleRate;
    uint32_t byteRate;
    uint16_t blockAlign;
    uint16_t bitsPerSample;
    char data[4];
    uint32_t dataSize;
};
#pragma pack(pop)

static bool WriteWavPcm(const std::wstring& path, const std::vector<uint8_t>& pcm,
                        int sampleRate, int channels, int bitsPerSample) {
    if (sampleRate <= 0 || channels <= 0 || (bitsPerSample != 8 && bitsPerSample != 16)) return false;

    WavHeader h = {};
    std::memcpy(h.riff, "RIFF", 4);
    std::memcpy(h.wave, "WAVE", 4);
    std::memcpy(h.fmt_, "fmt ", 4);
    std::memcpy(h.data, "data", 4);
    h.fmtSize = 16;
    h.audioFormat = 1; // PCM
    h.numChannels = (uint16_t)channels;
    h.sampleRate = (uint32_t)sampleRate;
    h.bitsPerSample = (uint16_t)bitsPerSample;
    h.blockAlign = (uint16_t)(channels * (bitsPerSample / 8));
    h.byteRate = (uint32_t)(sampleRate * (uint32_t)h.blockAlign);
    h.dataSize = (uint32_t)pcm.size();
    h.riffSize = 36 + h.dataSize;

    std::ofstream f(path, std::ios::binary);
    if (!f) return false;
    f.write(reinterpret_cast<const char*>(&h), sizeof(h));
    if (!pcm.empty()) f.write(reinterpret_cast<const char*>(pcm.data()), (std::streamsize)pcm.size());
    return (bool)f;
}

static int16_t ClampS16(int v) {
    if (v < -32768) return -32768;
    if (v > 32767) return 32767;
    return (int16_t)v;
}

static std::vector<int16_t> DecodeToMonoS16(const std::vector<uint8_t>& pcm, int channels, int bitsPerSample) {
    std::vector<int16_t> out;
    if (channels <= 0) return out;
    if (bitsPerSample == 16) {
        const size_t frameBytes = (size_t)channels * 2;
        const size_t frameCount = frameBytes ? (pcm.size() / frameBytes) : 0;
        out.reserve(frameCount);
        for (size_t i = 0; i < frameCount; i++) {
            int sum = 0;
            for (int ch = 0; ch < channels; ch++) {
                const size_t off = i * frameBytes + (size_t)ch * 2;
                int16_t s = (int16_t)(pcm[off] | (pcm[off + 1] << 8));
                sum += (int)s;
            }
            out.push_back(ClampS16(sum / channels));
        }
        return out;
    }
    if (bitsPerSample == 8) {
        const size_t frameBytes = (size_t)channels;
        const size_t frameCount = frameBytes ? (pcm.size() / frameBytes) : 0;
        out.reserve(frameCount);
        for (size_t i = 0; i < frameCount; i++) {
            int sum = 0;
            for (int ch = 0; ch < channels; ch++) {
                const size_t off = i * frameBytes + (size_t)ch;
                const int u = (int)pcm[off];
                const int s16 = (u - 128) << 8;
                sum += s16;
            }
            out.push_back(ClampS16(sum / channels));
        }
        return out;
    }
    return out;
}

static std::vector<int16_t> ResampleLinear(const std::vector<int16_t>& in, int inRate, int outRate) {
    std::vector<int16_t> out;
    if (in.empty() || inRate <= 0 || outRate <= 0) return out;
    if (inRate == outRate) return in;

    const double ratio = (double)inRate / (double)outRate;
    const size_t outCount = (size_t)((double)in.size() * (double)outRate / (double)inRate);
    out.resize(outCount);

    for (size_t i = 0; i < outCount; i++) {
        const double src = (double)i * ratio;
        size_t idx = (size_t)src;
        double frac = src - (double)idx;
        if (idx >= in.size()) idx = in.size() - 1;
        const int16_t s0 = in[idx];
        const int16_t s1 = (idx + 1 < in.size()) ? in[idx + 1] : s0;
        const double v = (double)s0 + ((double)s1 - (double)s0) * frac;
        out[i] = ClampS16((int)std::llround(v));
    }
    return out;
}

static std::vector<uint8_t> EncodeMonoS16ToBytes(const std::vector<int16_t>& mono) {
    std::vector<uint8_t> out;
    out.resize(mono.size() * 2);
    for (size_t i = 0; i < mono.size(); i++) {
        const int16_t s = mono[i];
        out[i * 2] = (uint8_t)(s & 0xFF);
        out[i * 2 + 1] = (uint8_t)((s >> 8) & 0xFF);
    }
    return out;
}

// -----------------------------------------------------------------------------
// WaveOut streaming player
// -----------------------------------------------------------------------------

class WaveOutPlayer {
public:
    WaveOutPlayer() = default;
    ~WaveOutPlayer() { close(); }

    bool open(int sampleRate, int channels, int bitsPerSample) {
        close();
        if (sampleRate <= 0 || channels <= 0) return false;
        if (bitsPerSample != 8 && bitsPerSample != 16) return false;

        WAVEFORMATEX wfx = {};
        wfx.wFormatTag = WAVE_FORMAT_PCM;
        wfx.nChannels = (WORD)channels;
        wfx.nSamplesPerSec = (DWORD)sampleRate;
        wfx.wBitsPerSample = (WORD)bitsPerSample;
        wfx.nBlockAlign = (WORD)(channels * (bitsPerSample / 8));
        wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * wfx.nBlockAlign;

        _drainedEvent = CreateEventW(nullptr, TRUE, TRUE, nullptr);
        if (!_drainedEvent) return false;

        MMRESULT mm = waveOutOpen(&_hwo, WAVE_MAPPER, &wfx, (DWORD_PTR)&WaveOutProc, (DWORD_PTR)this, CALLBACK_FUNCTION);
        if (mm != MMSYSERR_NOERROR) {
            CloseHandle(_drainedEvent);
            _drainedEvent = nullptr;
            _hwo = nullptr;
            return false;
        }
        _pending.store(0);
        return true;
    }

    void close() {
        if (_hwo) {
            waveOutReset(_hwo);
            waitDrained(2000);
            waveOutClose(_hwo);
            _hwo = nullptr;
        }
        if (_drainedEvent) {
            CloseHandle(_drainedEvent);
            _drainedEvent = nullptr;
        }
    }

    void stopNow() {
        if (_hwo) {
            waveOutReset(_hwo);
        }
    }

    bool feed(const uint8_t* data, size_t bytes) {
        if (!_hwo || !data || bytes == 0) return true;

        auto* buf = new Buffer();
        buf->data.assign(data, data + bytes);
        std::memset(&buf->hdr, 0, sizeof(buf->hdr));
        buf->hdr.lpData = reinterpret_cast<LPSTR>(buf->data.data());
        buf->hdr.dwBufferLength = (DWORD)buf->data.size();

        ResetEvent(_drainedEvent);
        _pending.fetch_add(1);

        MMRESULT mm = waveOutPrepareHeader(_hwo, &buf->hdr, sizeof(buf->hdr));
        if (mm != MMSYSERR_NOERROR) {
            _pending.fetch_sub(1);
            delete buf;
            if (_pending.load() == 0 && _drainedEvent) SetEvent(_drainedEvent);
            return false;
        }

        mm = waveOutWrite(_hwo, &buf->hdr, sizeof(buf->hdr));
        if (mm != MMSYSERR_NOERROR) {
            waveOutUnprepareHeader(_hwo, &buf->hdr, sizeof(buf->hdr));
            _pending.fetch_sub(1);
            delete buf;
            if (_pending.load() == 0 && _drainedEvent) SetEvent(_drainedEvent);
            return false;
        }

        // Ownership transfers to callback; it will delete.
        return true;
    }

    void waitDrained(DWORD timeoutMs) {
        if (!_drainedEvent) return;
        WaitForSingleObject(_drainedEvent, timeoutMs);
    }

private:
    struct Buffer {
        WAVEHDR hdr;
        std::vector<uint8_t> data;
    };

    static void CALLBACK WaveOutProc(HWAVEOUT /*hwo*/, UINT msg, DWORD_PTR instance, DWORD_PTR param1, DWORD_PTR /*param2*/) {
        if (msg != WOM_DONE) return;
        auto* self = reinterpret_cast<WaveOutPlayer*>(instance);
        if (!self) return;
        auto* hdr = reinterpret_cast<WAVEHDR*>(param1);
        if (!hdr) return;
        auto* buf = reinterpret_cast<Buffer*>(hdr);

        if (self->_hwo) {
            waveOutUnprepareHeader(self->_hwo, &buf->hdr, sizeof(buf->hdr));
        }

        long left = self->_pending.fetch_sub(1) - 1;
        if (left <= 0 && self->_drainedEvent) {
            SetEvent(self->_drainedEvent);
        }
        delete buf;
    }

    HWAVEOUT _hwo = nullptr;
    HANDLE _drainedEvent = nullptr;
    std::atomic<long> _pending{ 0 };
};

// -----------------------------------------------------------------------------
// Wrapper dynamic loader
// -----------------------------------------------------------------------------

struct WrapperApi {
    HMODULE dll = nullptr;
    void* handle = nullptr;

    using sv_initW_t = void* (__cdecl*)(const wchar_t* baseDllPath, int initialVoice);
    using sv_free_t = void(__cdecl*)(void*);
    using sv_stop_t = void(__cdecl*)(void*);
    using sv_startSpeakW_t = int(__cdecl*)(void*, const wchar_t*);
    using sv_read_t = int(__cdecl*)(void*, int* outType, int* outValue, uint8_t* outAudio, int outCap);

    using sv_set2_t = void(__cdecl*)(void*, int);
    using sv_get1_t = int(__cdecl*)(void*);

    using sv_getFormat_t = int(__cdecl*)(void*, int* sampleRate, int* channels, int* bits);
    using sv_setLead_t = void(__cdecl*)(void*, int);

    sv_initW_t sv_initW = nullptr;
    sv_free_t sv_free = nullptr;
    sv_stop_t sv_stop = nullptr;
    sv_startSpeakW_t sv_startSpeakW = nullptr;
    sv_read_t sv_read = nullptr;

    sv_set2_t sv_setRate = nullptr;
    sv_set2_t sv_setPitch = nullptr;
    sv_set2_t sv_setF0Range = nullptr;
    sv_set2_t sv_setF0Perturb = nullptr;
    sv_set2_t sv_setVowelFactor = nullptr;
    sv_set2_t sv_setAVBias = nullptr;
    sv_set2_t sv_setAFBias = nullptr;
    sv_set2_t sv_setAHBias = nullptr;
    sv_set2_t sv_setPersonality = nullptr;
    sv_set2_t sv_setF0Style = nullptr;
    sv_set2_t sv_setVoicingMode = nullptr;
    sv_set2_t sv_setGender = nullptr;
    sv_set2_t sv_setGlottalSource = nullptr;
    sv_set2_t sv_setSpeakingMode = nullptr;
    sv_set2_t sv_setVoice = nullptr;

    // Optional wrapper-only tuning
    sv_set2_t sv_setPauseFactor = nullptr;
    sv_set2_t sv_setTrimSilence = nullptr;
    sv_setLead_t sv_setMaxLeadMs = nullptr;

    sv_getFormat_t sv_getFormat = nullptr;

    bool loadFrom(const std::wstring& dllPath) {
        unload();
        dll = LoadLibraryW(dllPath.c_str());
        if (!dll) return false;

        auto gp = [&](auto& fn, const char* name) -> bool {
            fn = reinterpret_cast<std::decay_t<decltype(fn)>>(GetProcAddress(dll, name));
            return fn != nullptr;
        };

        // Required
        if (!gp(sv_initW, "sv_initW")) return false;
        if (!gp(sv_free, "sv_free")) return false;
        if (!gp(sv_stop, "sv_stop")) return false;
        if (!gp(sv_startSpeakW, "sv_startSpeakW")) return false;
        if (!gp(sv_read, "sv_read")) return false;

        // Settings
        gp(sv_setRate, "sv_setRate");
        gp(sv_setPitch, "sv_setPitch");
        gp(sv_setF0Range, "sv_setF0Range");
        gp(sv_setF0Perturb, "sv_setF0Perturb");
        gp(sv_setVowelFactor, "sv_setVowelFactor");
        gp(sv_setAVBias, "sv_setAVBias");
        gp(sv_setAFBias, "sv_setAFBias");
        gp(sv_setAHBias, "sv_setAHBias");
        gp(sv_setPersonality, "sv_setPersonality");
        gp(sv_setF0Style, "sv_setF0Style");
        gp(sv_setVoicingMode, "sv_setVoicingMode");
        gp(sv_setGender, "sv_setGender");
        gp(sv_setGlottalSource, "sv_setGlottalSource");
        gp(sv_setSpeakingMode, "sv_setSpeakingMode");
        gp(sv_setVoice, "sv_setVoice");

        // Optional
        gp(sv_setPauseFactor, "sv_setPauseFactor");
        gp(sv_setTrimSilence, "sv_setTrimSilence");
        gp(sv_setMaxLeadMs, "sv_setMaxLeadMs");
        gp(sv_getFormat, "sv_getFormat");
        return true;
    }

    void unload() {
        if (handle && sv_free) {
            sv_free(handle);
            handle = nullptr;
        }
        if (dll) {
            FreeLibrary(dll);
            dll = nullptr;
        }
    }
};

// -----------------------------------------------------------------------------
// App state
// -----------------------------------------------------------------------------

enum class JobMode {
    Speak,
    SaveWav,
};

struct UiSettings {
    int voice = 1;       // 1=English, 2=Spanish
    int variant = 0;     // personality

    int ratePct = 50;
    int pitchPct = 4;
    int inflectionPct = 25;
    int perturbPct = 0;
    int vfactorPct = 20;
    int avbiasPct = 50;
    int afbiasPct = 50;
    int ahbiasPct = 50;
    int pausePct = 50;

    int intstyle = 0;
    int vmode = 0;
    int gender = 0;
    int glot = 0;
    int smode = 0;
};

struct AppState {
    HWND dlg = nullptr;
    WrapperApi api;

    std::mutex statusMtx;
    std::wstring pendingStatus;

    std::atomic<bool> jobRunning{ false };
    std::atomic<bool> cancelRequested{ false };

    std::thread worker;

    UiSettings lastApplied;
    bool lastAppliedValid = false;

    bool initializing = false; // suppress "touched" flags while building the dialog
    
    // Explicit-setting flags (session-only).
    //
    // SoftVoice personalities (especially the fun ones like Robotoid/Martian) come with their own
    // internal timbre defaults. If we blindly push our UI defaults into the engine, we overwrite
    // those and the personality sounds wrong (e.g. whispery instead of robotic).
    //
    // Rule here matches the NVDA driver strategy:
    //   - Rate/Pitch are always applied.
    //   - For Variant != 0, only apply timbre/style knobs if the user has explicitly changed them.
    bool expInflection = false;
    bool expPerturb    = false;
    bool expVfactor    = false;
    bool expAvbias     = false;
    bool expAfbias     = false;
    bool expAhbias     = false;

    bool expIntstyle   = false;
    bool expVmode      = false;
    bool expGender     = false;
    bool expGlot       = false;
};

static constexpr UINT WM_APP_STATUS = WM_APP + 1;
static constexpr UINT WM_APP_DONE = WM_APP + 2;

// -----------------------------------------------------------------------------
// UI -> settings
// -----------------------------------------------------------------------------

static int ComboGetItemDataInt(HWND combo) {
    LRESULT sel = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR) return 0;
    LRESULT data = SendMessageW(combo, CB_GETITEMDATA, (WPARAM)sel, 0);
    return (int)data;
}

static UiSettings ReadSettingsFromUi(HWND dlg) {
    UiSettings s;
    s.voice = ComboGetItemDataInt(GetDlgItem(dlg, IDC_VOICE));
    s.variant = ComboGetItemDataInt(GetDlgItem(dlg, IDC_VARIANT));
    s.smode = ComboGetItemDataInt(GetDlgItem(dlg, IDC_SMODE));
    s.intstyle = ComboGetItemDataInt(GetDlgItem(dlg, IDC_INTSTYLE));
    s.vmode = ComboGetItemDataInt(GetDlgItem(dlg, IDC_VMODE));
    s.gender = ComboGetItemDataInt(GetDlgItem(dlg, IDC_GENDER));
    s.glot = ComboGetItemDataInt(GetDlgItem(dlg, IDC_GLOT));

    BOOL ok = FALSE;
    s.ratePct = ClampInt((int)GetDlgItemInt(dlg, IDC_RATE, &ok, FALSE), 0, 100);
    s.pitchPct = ClampInt((int)GetDlgItemInt(dlg, IDC_PITCH, &ok, FALSE), 0, 100);
    s.inflectionPct = ClampInt((int)GetDlgItemInt(dlg, IDC_INFLECTION, &ok, FALSE), 0, 100);
    s.pausePct = ClampInt((int)GetDlgItemInt(dlg, IDC_PAUSE, &ok, FALSE), 0, 100);

    s.perturbPct = ClampInt((int)GetDlgItemInt(dlg, IDC_PERTURB, &ok, FALSE), 0, 100);
    s.vfactorPct = ClampInt((int)GetDlgItemInt(dlg, IDC_VFACTOR, &ok, FALSE), 0, 100);
    s.avbiasPct = ClampInt((int)GetDlgItemInt(dlg, IDC_AVBIAS, &ok, FALSE), 0, 100);
    s.afbiasPct = ClampInt((int)GetDlgItemInt(dlg, IDC_AFBIAS, &ok, FALSE), 0, 100);
    s.ahbiasPct = ClampInt((int)GetDlgItemInt(dlg, IDC_AHBIAS, &ok, FALSE), 0, 100);

    return s;
}

static std::wstring ReadTextBox(HWND dlg) {
    HWND edit = GetDlgItem(dlg, IDC_TEXT);
    int len = GetWindowTextLengthW(edit);
    if (len <= 0) return L"";
    std::wstring s;
    s.resize((size_t)len);
    GetWindowTextW(edit, s.data(), len + 1);
    return s;
}

static std::wstring ReadEnginePath(HWND dlg) {
    wchar_t buf[MAX_PATH] = {};
    GetDlgItemTextW(dlg, IDC_ENGINE_PATH, buf, (int)_countof(buf));
    return Trim(std::wstring(buf));
}

// -----------------------------------------------------------------------------
// Settings -> wrapper
// -----------------------------------------------------------------------------

static bool ApplyWrapperSettings(AppState* st, const UiSettings& s, const std::wstring& tibasePath) {
    if (!st || !st->api.handle) return false;
    void* h = st->api.handle;
    (void)tibasePath;

    const bool firstApply = !st->lastAppliedValid;
    const bool voiceChanged = firstApply || (st->lastApplied.voice != s.voice);
    const bool variantChanged = firstApply || (st->lastApplied.variant != s.variant);
    const bool presetChanged = voiceChanged || variantChanged;

    auto changed = [&](auto getter) -> bool {
        if (firstApply) return true;
        return getter(st->lastApplied) != getter(s);
    };
    auto changedOrPreset = [&](auto getter) -> bool {
        if (presetChanged) return true;
        return changed(getter);
    };

    // 1) Voice (language)
    if (st->api.sv_setVoice && changed([](const UiSettings& x) { return x.voice; })) {
        st->api.sv_setVoice(h, s.voice);
    }

    // 2) Personality (variant)
    //
    // IMPORTANT: keep the "wake-up" behavior for Variant 0 (Male) from the known-good build.
    // Some SoftVoice installs won't actually synthesize Male unless we poke the personality state.
    if (st->api.sv_setPersonality && variantChanged) {
        if (s.variant == 0) {
            // Wake-up toggle: 0 -> 1 -> 0.
            st->api.sv_setPersonality(h, 1);
            Sleep(20);
        }
        st->api.sv_setPersonality(h, s.variant);
    }

    // 3) Speaking mode
    //
    // We implement word/spell modes in our own text splitter. Keep the engine in "Natural".
    if (firstApply && st->api.sv_setSpeakingMode) {
        st->api.sv_setSpeakingMode(h, 0);
    }

    // 4) Always-safe knobs: Rate + Pitch
    //
    // Personality switches can reset internals, so we re-assert these after a preset change.
    if (st->api.sv_setRate && changedOrPreset([](const UiSettings& x) { return x.ratePct; })) {
        st->api.sv_setRate(h, PercentToParam(s.ratePct, 20, 500));
    }
    if (st->api.sv_setPitch && changedOrPreset([](const UiSettings& x) { return x.pitchPct; })) {
        st->api.sv_setPitch(h, PercentToParam(s.pitchPct, 10, 2000));
    }

    const bool isMale = (s.variant == 0);

    // 5) Timbre + style knobs
    //
    // The important trick for the fun personalities (Robotoid/Martian/etc):
    // they come with their own internal defaults for perturb/biases/voicing/etc.
    // If we blindly push our UI defaults, we overwrite those and the voice sounds wrong.
    //
    // Strategy (mirrors the NVDA driver):
    //   - For Male (variant 0): do NOT push these knobs by default (male is fragile).
    //   - For other variants: only push a knob if the user explicitly touched it.
    if (!isMale) {
        if (st->expInflection && st->api.sv_setF0Range && changedOrPreset([](const UiSettings& x) { return x.inflectionPct; })) {
            st->api.sv_setF0Range(h, PercentToParam(s.inflectionPct, 0, 500));
        }
        if (st->expPerturb && st->api.sv_setF0Perturb && changedOrPreset([](const UiSettings& x) { return x.perturbPct; })) {
            st->api.sv_setF0Perturb(h, PercentToParam(s.perturbPct, 0, 500));
        }
        if (st->expVfactor && st->api.sv_setVowelFactor && changedOrPreset([](const UiSettings& x) { return x.vfactorPct; })) {
            st->api.sv_setVowelFactor(h, PercentToParam(s.vfactorPct, 0, 500));
        }

        if (st->expAvbias && st->api.sv_setAVBias && changedOrPreset([](const UiSettings& x) { return x.avbiasPct; })) {
            st->api.sv_setAVBias(h, PercentToParam(s.avbiasPct, -50, 50));
        }
        if (st->expAfbias && st->api.sv_setAFBias && changedOrPreset([](const UiSettings& x) { return x.afbiasPct; })) {
            st->api.sv_setAFBias(h, PercentToParam(s.afbiasPct, -50, 50));
        }
        if (st->expAhbias && st->api.sv_setAHBias && changedOrPreset([](const UiSettings& x) { return x.ahbiasPct; })) {
            st->api.sv_setAHBias(h, PercentToParam(s.ahbiasPct, -50, 50));
        }

        // Enums (only if explicitly changed in the UI)
        if (st->expIntstyle && st->api.sv_setF0Style && changedOrPreset([](const UiSettings& x) { return x.intstyle; })) {
            st->api.sv_setF0Style(h, s.intstyle);
        }
        if (st->expVmode && st->api.sv_setVoicingMode && changedOrPreset([](const UiSettings& x) { return x.vmode; })) {
            st->api.sv_setVoicingMode(h, s.vmode);
        }
        if (st->expGender && st->api.sv_setGender && changedOrPreset([](const UiSettings& x) { return x.gender; })) {
            st->api.sv_setGender(h, s.gender);
        }
        if (st->expGlot && st->api.sv_setGlottalSource && changedOrPreset([](const UiSettings& x) { return x.glot; })) {
            st->api.sv_setGlottalSource(h, s.glot);
        }
    }

    // Wrapper-only knobs (safe)
    if (st->api.sv_setPauseFactor && changedOrPreset([](const UiSettings& x) { return x.pausePct; })) {
        st->api.sv_setPauseFactor(h, 100 - s.pausePct);
        if (st->api.sv_setTrimSilence) {
            st->api.sv_setTrimSilence(h, (s.pausePct < 50) ? 1 : 0);
        }
    }

    st->lastApplied = s;
    st->lastAppliedValid = true;
    return true;
}

// -----------------------------------------------------------------------------
// Status posting
// -----------------------------------------------------------------------------

static void PostStatus(AppState* st, const std::wstring& msg) {
    if (!st || !st->dlg) return;
    {
        std::lock_guard<std::mutex> g(st->statusMtx);
        st->pendingStatus = msg;
    }
    PostMessageW(st->dlg, WM_APP_STATUS, 0, 0);
}

// -----------------------------------------------------------------------------
// Synthesis worker
// -----------------------------------------------------------------------------

static bool EnsureWrapperReady(AppState* st, const std::wstring& wrapperDllPath, const std::wstring& tibasePath) {
    if (!st) return false;

    if (!st->api.dll) {
        if (!st->api.loadFrom(wrapperDllPath)) {
            return false;
        }
    }

    if (!st->api.handle) {
        st->api.handle = st->api.sv_initW(tibasePath.c_str(), 1);
        if (!st->api.handle) return false;
    }
    return true;
}

static bool PumpOneSegment(AppState* st, JobMode mode, const std::wstring& seg,
                           WaveOutPlayer* player,
                           std::vector<uint8_t>* rawAudio,
                           int* ioSampleRate, int* ioChannels, int* ioBits) {
    if (!st || !st->api.handle) return false;
    if (seg.empty()) return true;
    if (st->cancelRequested.load()) return false;

    st->api.sv_startSpeakW(st->api.handle, seg.c_str());

    std::vector<uint8_t> buf;
    buf.resize(32768);

    bool playerOpened = (player != nullptr && *ioSampleRate > 0);

    while (!st->cancelRequested.load()) {
        int t = SV_ITEM_NONE;
        int v = 0;
        int n = st->api.sv_read(st->api.handle, &t, &v, buf.data(), (int)buf.size());

        if (t == SV_ITEM_AUDIO && n > 0) {
            if (st->api.sv_getFormat && (*ioSampleRate <= 0 || *ioChannels <= 0 || *ioBits <= 0)) {
                int sr = 0, ch = 0, bits = 0;
                if (st->api.sv_getFormat(st->api.handle, &sr, &ch, &bits) == 1) {
                    *ioSampleRate = sr;
                    *ioChannels = ch;
                    *ioBits = bits;
                }
            }

            if (mode == JobMode::Speak) {
                if (player && !playerOpened) {
                    int sr = (*ioSampleRate > 0) ? *ioSampleRate : TARGET_WAV_RATE;
                    int ch = (*ioChannels > 0) ? *ioChannels : 1;
                    int bits = (*ioBits > 0) ? *ioBits : 16;
                    playerOpened = player->open(sr, ch, bits);
                }
                if (playerOpened && player) {
                    player->feed(buf.data(), (size_t)n);
                }
            } else if (mode == JobMode::SaveWav) {
                if (rawAudio) {
                    rawAudio->insert(rawAudio->end(), buf.begin(), buf.begin() + n);
                }
            }
            continue;
        }

        if (t == SV_ITEM_DONE) {
            return true;
        }
        if (t == SV_ITEM_ERROR) {
            wchar_t tmp[128] = {};
            swprintf_s(tmp, L"SoftVoice error (%d)", v);
            PostStatus(st, tmp);
            return false;
        }

        // No data yet.
        Sleep(1);
    }
    return false;
}

static void WorkerThread(AppState* st, JobMode mode, std::wstring wavOutPath) {
    if (!st) return;

    const std::wstring exeDir = GetExeDir();
    const std::wstring wrapperDllPath = exeDir.empty() ? L"softvoice_wrapper.dll" : (exeDir + L"\\softvoice_wrapper.dll");
    const std::wstring tibasePath = ReadEnginePath(st->dlg);

    if (tibasePath.empty() || !FileExists(tibasePath) || !IsTibase32Path(tibasePath)) {
        PostStatus(st, L"Please select tibase32.dll (exact file name).");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }
    if (!FileExists(wrapperDllPath)) {
        PostStatus(st, L"softvoice_wrapper.dll not found next to the app.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    if (!EnsureWrapperReady(st, wrapperDllPath, tibasePath)) {
        PostStatus(st, L"Wrapper init failed.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    // Snapshot settings + text
    const UiSettings settings = ReadSettingsFromUi(st->dlg);
    const std::wstring fullText = ReadTextBox(st->dlg);
    if (Trim(fullText).empty()) {
        PostStatus(st, L"Nothing to speak.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    st->cancelRequested.store(false);

    // Stop any previous audio first (before changing personality/style).
    if (st->api.handle && st->api.sv_stop) {
        st->api.sv_stop(st->api.handle);
    }

    if (!ApplyWrapperSettings(st, settings, tibasePath)) {
        PostStatus(st, L"Failed to apply settings.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }


    std::vector<std::wstring> segments = SplitForSoftVoice(fullText, settings.smode);
    if (segments.empty()) {
        PostStatus(st, L"Nothing to speak.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    int inRate = 0, inCh = 0, inBits = 0;

    if (mode == JobMode::Speak) {
        PostStatus(st, L"Speaking...");
        WaveOutPlayer player;
        for (const auto& seg : segments) {
            if (st->cancelRequested.load()) break;
            if (!PumpOneSegment(st, mode, seg, &player, nullptr, &inRate, &inCh, &inBits)) break;
        }
        if (st->cancelRequested.load()) {
            player.stopNow();
            PostStatus(st, L"Stopped.");
        } else {
            player.waitDrained(5000);
            PostStatus(st, L"Done.");
        }
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    // Save WAV
    PostStatus(st, L"Rendering WAV...");
    std::vector<uint8_t> raw;
    raw.reserve(1024 * 256);
    for (const auto& seg : segments) {
        if (st->cancelRequested.load()) break;
        if (!PumpOneSegment(st, mode, seg, nullptr, &raw, &inRate, &inCh, &inBits)) break;
    }

    if (st->cancelRequested.load()) {
        PostStatus(st, L"Stopped.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    if (inRate <= 0) inRate = TARGET_WAV_RATE;
    if (inCh <= 0) inCh = 1;
    if (inBits <= 0) inBits = 16;

    // Convert to 11025 Hz, mono, 16-bit.
    std::vector<int16_t> mono = DecodeToMonoS16(raw, inCh, inBits);
    std::vector<int16_t> res = ResampleLinear(mono, inRate, TARGET_WAV_RATE);
    std::vector<uint8_t> pcm = EncodeMonoS16ToBytes(res);

    if (!WriteWavPcm(wavOutPath, pcm, TARGET_WAV_RATE, TARGET_WAV_CHANNELS, TARGET_WAV_BITS)) {
        PostStatus(st, L"Failed to write WAV file.");
        PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
        return;
    }

    PostStatus(st, L"Saved WAV.");
    PostMessageW(st->dlg, WM_APP_DONE, 0, 0);
}

// -----------------------------------------------------------------------------
// Dialog helpers
// -----------------------------------------------------------------------------

static void SetButtonsEnabled(HWND dlg, bool idle) {
    EnableWindow(GetDlgItem(dlg, IDC_SPEAK), idle);
    EnableWindow(GetDlgItem(dlg, IDC_SAVE_WAV), idle);
    EnableWindow(GetDlgItem(dlg, IDC_OPEN_TEXT), idle);
    EnableWindow(GetDlgItem(dlg, IDC_ENGINE_BROWSE), idle);
    EnableWindow(GetDlgItem(dlg, IDC_STOP), !idle);
}

static void InitSpin(HWND dlg, int spinId, int minV, int maxV) {
    HWND sp = GetDlgItem(dlg, spinId);
    if (!sp) return;
    SendMessageW(sp, UDM_SETRANGE32, (WPARAM)minV, (LPARAM)maxV);
}

static void ComboAdd(HWND combo, int id, const wchar_t* label) {
    LRESULT idx = SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)label);
    if (idx != CB_ERR && idx != CB_ERRSPACE) {
        SendMessageW(combo, CB_SETITEMDATA, (WPARAM)idx, (LPARAM)id);
    }
}

static void ComboSelectByData(HWND combo, int value) {
    const LRESULT count = SendMessageW(combo, CB_GETCOUNT, 0, 0);
    for (LRESULT i = 0; i < count; i++) {
        LRESULT data = SendMessageW(combo, CB_GETITEMDATA, (WPARAM)i, 0);
        if ((int)data == value) {
            SendMessageW(combo, CB_SETCURSEL, (WPARAM)i, 0);
            return;
        }
    }
    SendMessageW(combo, CB_SETCURSEL, 0, 0);
}

static void LoadTextFileIntoEdit(HWND dlg, const std::wstring& path) {
    std::vector<uint8_t> bytes;
    if (!ReadWholeFileBytes(path, bytes)) {
        MessageBoxW(dlg, L"Could not read the file.", L"SoftVoice Speak", MB_ICONERROR);
        return;
    }
    std::wstring w = BytesToWideBestEffort(bytes);
    if (w.empty() && !bytes.empty()) {
        MessageBoxW(dlg, L"The file could not be decoded as text.", L"SoftVoice Speak", MB_ICONERROR);
        return;
    }
    SetDlgItemTextW(dlg, IDC_TEXT, w.c_str());
}

// -----------------------------------------------------------------------------
// Dialog proc
// -----------------------------------------------------------------------------

static INT_PTR CALLBACK MainDlgProc(HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) {
    AppState* st = reinterpret_cast<AppState*>(GetWindowLongPtrW(dlg, GWLP_USERDATA));

    switch (msg) {
    case WM_INITDIALOG: {
        auto* state = reinterpret_cast<AppState*>(lParam);
        state->dlg = dlg;
        SetWindowLongPtrW(dlg, GWLP_USERDATA, (LONG_PTR)state);
        st = state;
        st->initializing = true;

        InitCommonControls();

        // Default engine path: tibase32.dll next to exe.
        const std::wstring exeDir = GetExeDir();
        const std::wstring tibaseGuess = exeDir.empty() ? L"tibase32.dll" : (exeDir + L"\\tibase32.dll");
        if (FileExists(tibaseGuess)) {
            SetDlgItemTextW(dlg, IDC_ENGINE_PATH, tibaseGuess.c_str());
        }

        // Spin ranges.
        InitSpin(dlg, IDC_RATE_SPIN, 0, 100);
        InitSpin(dlg, IDC_PITCH_SPIN, 0, 100);
        InitSpin(dlg, IDC_INFLECTION_SPIN, 0, 100);
        InitSpin(dlg, IDC_PAUSE_SPIN, 0, 100);
        InitSpin(dlg, IDC_PERTURB_SPIN, 0, 100);
        InitSpin(dlg, IDC_VFACTOR_SPIN, 0, 100);
        InitSpin(dlg, IDC_AVBIAS_SPIN, 0, 100);
        InitSpin(dlg, IDC_AFBIAS_SPIN, 0, 100);
        InitSpin(dlg, IDC_AHBIAS_SPIN, 0, 100);

        // Default numeric values (match NVDA driver defaults).
        SetDlgItemInt(dlg, IDC_RATE, 50, FALSE);
        SetDlgItemInt(dlg, IDC_PITCH, 4, FALSE);
        SetDlgItemInt(dlg, IDC_INFLECTION, 25, FALSE);
        SetDlgItemInt(dlg, IDC_PAUSE, 50, FALSE);
        SetDlgItemInt(dlg, IDC_PERTURB, 0, FALSE);
        SetDlgItemInt(dlg, IDC_VFACTOR, 20, FALSE);
        SetDlgItemInt(dlg, IDC_AVBIAS, 50, FALSE);
        SetDlgItemInt(dlg, IDC_AFBIAS, 50, FALSE);
        SetDlgItemInt(dlg, IDC_AHBIAS, 50, FALSE);

        // Fill combos.
        HWND voice = GetDlgItem(dlg, IDC_VOICE);
        ComboAdd(voice, 1, L"English");
        ComboAdd(voice, 2, L"Spanish");
        ComboSelectByData(voice, 1);

        HWND variant = GetDlgItem(dlg, IDC_VARIANT);
        // Personality list copied from the NVDA driver.
        ComboAdd(variant, 0, L"Male");
        ComboAdd(variant, 1, L"Female");
        ComboAdd(variant, 2, L"Large Male");
        ComboAdd(variant, 3, L"Child");
        ComboAdd(variant, 4, L"Giant Male");
        ComboAdd(variant, 5, L"Mellow Female");
        ComboAdd(variant, 6, L"Mellow Male");
        ComboAdd(variant, 7, L"Crisp Male");
        ComboAdd(variant, 8, L"The Fly");
        ComboAdd(variant, 9, L"Robotoid");
        ComboAdd(variant, 10, L"Martian");
        ComboAdd(variant, 11, L"Colossus");
        ComboAdd(variant, 12, L"Fast Fred");
        ComboAdd(variant, 13, L"Old Woman");
        ComboAdd(variant, 14, L"Munchkin");
        ComboAdd(variant, 15, L"Troll");
        ComboAdd(variant, 16, L"Nerd");
        ComboAdd(variant, 17, L"Milktoast");
        ComboAdd(variant, 18, L"Tipsy");
        ComboAdd(variant, 19, L"Choirboy");
        ComboSelectByData(variant, 0);

        HWND smode = GetDlgItem(dlg, IDC_SMODE);
        ComboAdd(smode, 0, L"Natural");
        ComboAdd(smode, 1, L"Word-at-a-time");
        ComboAdd(smode, 2, L"Spelled");
        ComboSelectByData(smode, 0);

        HWND intstyle = GetDlgItem(dlg, IDC_INTSTYLE);
        ComboAdd(intstyle, 0, L"normal1");
        ComboAdd(intstyle, 1, L"normal2");
        ComboAdd(intstyle, 2, L"monotone");
        ComboAdd(intstyle, 3, L"sung");
        ComboAdd(intstyle, 4, L"random");
        ComboSelectByData(intstyle, 0);

        HWND vmode = GetDlgItem(dlg, IDC_VMODE);
        ComboAdd(vmode, 0, L"normal");
        ComboAdd(vmode, 1, L"breathy");
        ComboAdd(vmode, 2, L"whispered");
        ComboSelectByData(vmode, 0);

        HWND gender = GetDlgItem(dlg, IDC_GENDER);
        ComboAdd(gender, 0, L"male");
        ComboAdd(gender, 1, L"female");
        ComboAdd(gender, 2, L"child");
        ComboAdd(gender, 3, L"giant");
        ComboSelectByData(gender, 0);

        HWND glot = GetDlgItem(dlg, IDC_GLOT);
        ComboAdd(glot, 0, L"default");
        ComboAdd(glot, 1, L"male");
        ComboAdd(glot, 2, L"female");
        ComboAdd(glot, 3, L"child");
        ComboAdd(glot, 4, L"high");
        ComboAdd(glot, 5, L"mellow");
        ComboAdd(glot, 6, L"impulse");
        ComboAdd(glot, 7, L"odd");
        ComboAdd(glot, 8, L"colossus");
        ComboSelectByData(glot, 0);

        SetDlgItemTextW(dlg, IDC_STATUS, L"Ready.");
        SetButtonsEnabled(dlg, true);
        st->initializing = false;
        return TRUE;
    }

    case WM_COMMAND: {
        const int id = LOWORD(wParam);
        const int code = HIWORD(wParam);

        if (!st) return FALSE;

        // Track which "timbre/style" knobs the user explicitly touched.
        //
        // For non-Male personalities, we avoid pushing our default knob values into the engine
        // unless the user has changed them (otherwise we stomp the personality's own defaults).
        if (!st->initializing) {
            if (code == EN_CHANGE) {
                switch (id) {
                case IDC_INFLECTION: st->expInflection = true; break;
                case IDC_PERTURB:    st->expPerturb    = true; break;
                case IDC_VFACTOR:    st->expVfactor    = true; break;
                case IDC_AVBIAS:     st->expAvbias     = true; break;
                case IDC_AFBIAS:     st->expAfbias     = true; break;
                case IDC_AHBIAS:     st->expAhbias     = true; break;
                default: break;
                }
            } else if (code == CBN_SELCHANGE) {
                switch (id) {
                case IDC_INTSTYLE: st->expIntstyle = true; break;
                case IDC_VMODE:    st->expVmode    = true; break;
                case IDC_GENDER:   st->expGender   = true; break;
                case IDC_GLOT:     st->expGlot     = true; break;
                default: break;
                }
            }
        }


        switch (id) {
        case IDC_ENGINE_BROWSE: {
            const wchar_t* filter = L"SoftVoice engine (tibase32.dll)\0tibase32.dll\0\0";
            std::wstring p = BrowseForFile(dlg, false, L"Select tibase32.dll", filter, L"dll");
            if (!p.empty()) {
                if (!IsTibase32Path(p)) {
                    MessageBoxW(dlg, L"Please choose tibase32.dll (exact file name).", L"SoftVoice Speak", MB_OK | MB_ICONWARNING);
                } else {
                    SetDlgItemTextW(dlg, IDC_ENGINE_PATH, p.c_str());
                }
            }
            return TRUE;
        }
        case IDC_OPEN_TEXT: {
            const wchar_t* filter = L"Text files\0*.txt;*.text\0All files\0*.*\0\0";
            std::wstring p = BrowseForFile(dlg, false, L"Open text file", filter, nullptr);
            if (!p.empty()) LoadTextFileIntoEdit(dlg, p);
            return TRUE;
        }
        case IDC_SPEAK: {
            if (st->jobRunning.load()) return TRUE;
            st->jobRunning.store(true);
            st->cancelRequested.store(false);
            SetButtonsEnabled(dlg, false);
            {
                std::lock_guard<std::mutex> g(st->statusMtx);
                st->pendingStatus = L"Starting...";
            }
            SetDlgItemTextW(dlg, IDC_STATUS, L"Starting...");

            st->worker = std::thread(WorkerThread, st, JobMode::Speak, std::wstring());
            return TRUE;
        }
        case IDC_SAVE_WAV: {
            if (st->jobRunning.load()) return TRUE;
            const wchar_t* filter = L"WAV files\0*.wav\0All files\0*.*\0\0";
            std::wstring out = BrowseForFile(dlg, true, L"Save WAV", filter, L"wav");
            if (out.empty()) return TRUE;

            st->jobRunning.store(true);
            st->cancelRequested.store(false);
            SetButtonsEnabled(dlg, false);
            SetDlgItemTextW(dlg, IDC_STATUS, L"Starting...");
            st->worker = std::thread(WorkerThread, st, JobMode::SaveWav, out);
            return TRUE;
        }
        case IDC_STOP: {
            if (!st->jobRunning.load()) return TRUE;
            st->cancelRequested.store(true);
            if (st->api.handle && st->api.sv_stop) {
                st->api.sv_stop(st->api.handle);
            }
            SetDlgItemTextW(dlg, IDC_STATUS, L"Stopping...");
            return TRUE;
        }
        case IDCANCEL: {
            SendMessageW(dlg, WM_CLOSE, 0, 0);
            return TRUE;
        }
        default:
            break;
        }
        return FALSE;
    }

    case WM_APP_STATUS: {
        if (!st) return TRUE;
        std::wstring msg;
        {
            std::lock_guard<std::mutex> g(st->statusMtx);
            msg = st->pendingStatus;
        }
        SetDlgItemTextW(dlg, IDC_STATUS, msg.c_str());
        return TRUE;
    }

    case WM_APP_DONE: {
        if (!st) return TRUE;
        if (st->worker.joinable()) st->worker.join();
        st->jobRunning.store(false);
        st->cancelRequested.store(false);
        SetButtonsEnabled(dlg, true);
        return TRUE;
    }

    case WM_CLOSE: {
        if (!st) {
            EndDialog(dlg, 0);
            return TRUE;
        }

        if (st->jobRunning.load()) {
            st->cancelRequested.store(true);
            if (st->api.handle && st->api.sv_stop) st->api.sv_stop(st->api.handle);
            if (st->worker.joinable()) st->worker.join();
            st->jobRunning.store(false);
        }
        st->api.unload();
        EndDialog(dlg, 0);
        return TRUE;
    }
    }

    return FALSE;
}

// -----------------------------------------------------------------------------
// Entry
// -----------------------------------------------------------------------------

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int) {
    INITCOMMONCONTROLSEX icc = {};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_STANDARD_CLASSES | ICC_UPDOWN_CLASS;
    InitCommonControlsEx(&icc);

    AppState st;
    DialogBoxParamW(hInst, MAKEINTRESOURCEW(IDD_MAIN), nullptr, MainDlgProc, (LPARAM)&st);
    return 0;
}
