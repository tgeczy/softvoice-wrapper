# -*- coding: utf-8 -*-
# synthDrivers/sv.py
#
# SoftVoice (late 90s) NVDA synth driver.
#
# FINAL FIXES:
# 1. PAUSE FACTOR INVERTED:
#    - Slider 0 (User) -> DLL 100 (Engine) = Zero Pauses (Robot).
#    - Slider 100 (User) -> DLL 0 (Engine) = Full Natural Pauses.
#
# 2. SECRET WEAPON (sv_setTrimSilence):
#    - Automatically enables 'Trim Silence' when the slider is below 50 (low pauses).
#    - This forces the engine to cut startup latency for fast response.

import os
import ctypes
import threading
import queue
import time
import re
from collections import OrderedDict

import nvwave
import config

from logHandler import log
from synthDriverHandler import SynthDriver, VoiceInfo, synthDoneSpeaking, synthIndexReached
from speech.commands import IndexCommand
from autoSettingsUtils.driverSetting import DriverSetting, NumericDriverSetting

# --- Wrapper Constants ---
SV_ITEM_NONE = 0
SV_ITEM_AUDIO = 1
SV_ITEM_DONE = 2
SV_ITEM_ERROR = 3

# --- Global DLL Management ---
_wrapperLock = threading.RLock()
_wrapperHandle = None
_wrapperRefCount = 0
_wrapperDll = None

MAX_STRING_LENGTH = 1200

# --- Background Thread ---
class _BgThread(threading.Thread):
    def __init__(self, q: "queue.Queue", stopEvent: "threading.Event"):
        super().__init__(name=f"{self.__class__.__module__}.{self.__class__.__qualname__}")
        self.daemon = True
        self._q = q
        self._stop = stopEvent

    def run(self):
        while not self._stop.is_set():
            try:
                item = self._q.get(timeout=0.2)
            except queue.Empty:
                continue
            try:
                if item is None: return
                func, args, kwargs = item
                func(*args, **kwargs)
            except Exception:
                log.error("SoftVoice: error running background synth function", exc_info=True)
            finally:
                try: self._q.task_done()
                except Exception: pass

# --- Text Cleaning ---
_PUNCT_TRANSLATE = str.maketrans({
    "’": "'", "‘": "'", "“": '"', "”": '"', "–": "-", "—": "-", "…": "...", "\u00a0": " ",
})
_control_re = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]+")
_STRIP_CHARS = {"\ufeff", "\u00ad", "\u200b", "\u200c", "\u200d", "\u200e", "\u200f"}
_labelColonRe = re.compile(r"([A-Za-z]{2,})\s*:\s*([A-Za-z])")
_labelSemiRe = re.compile(r"([A-Za-z]{2,})\s*;\s*([A-Za-z])")
_spellWordRe = re.compile(r"[A-Za-z0-9]+")

def _sanitizeText(s: str) -> str:
    if not s: return ""
    s = s.translate(_PUNCT_TRANSLATE)
    for ch in _STRIP_CHARS: s = s.replace(ch, "")
    s = _control_re.sub(" ", s)
    s = "".join((c if ord(c) <= 0xFFFF else " ") for c in s)
    s = " ".join(s.split())
    return s.strip()

def _find_tibase32(base_path: str) -> str:
    for p in (os.path.join(base_path, "tibase32.dll"), os.path.join(base_path, "TIBASE32.DLL")):
        if os.path.isfile(p): return p
    return ""

def _canonicalizeWindowsPath(p: str) -> str:
    if not p: return ""
    try: p = os.path.abspath(p)
    except Exception: pass
    try: p = os.path.normpath(p)
    except Exception: pass
    return p.replace("/", "\\")

# --- Enum Definitions ---
variants = OrderedDict()
def _v(_id, label): variants[str(_id)] = VoiceInfo(str(_id), label)
_v(0, "Male"); _v(1, "Female"); _v(2, "Large Male"); _v(3, "Child"); _v(4, "Giant Male")
_v(5, "Mellow Female"); _v(6, "Mellow Male"); _v(7, "Crisp Male"); _v(8, "The Fly")
_v(9, "Robotoid"); _v(10, "Martian"); _v(11, "Colossus"); _v(12, "Fast Fred")
_v(13, "Old Woman"); _v(14, "Munchkin"); _v(15, "Troll"); _v(16, "Nerd")
_v(17, "Milktoast"); _v(18, "Tipsy"); _v(19, "Choirboy")

intstyles = OrderedDict()
def _i(_id, label): intstyles[str(_id)] = VoiceInfo(str(_id), label)
_i(0, "normal1"); _i(1, "normal2"); _i(2, "monotone"); _i(3, "sung"); _i(4, "random")

vmodes = OrderedDict()
def _m(_id, label): vmodes[str(_id)] = VoiceInfo(str(_id), label)
_m(0, "normal"); _m(1, "breathy"); _m(2, "whispered")

genders = OrderedDict()
def _g(_id, label): genders[str(_id)] = VoiceInfo(str(_id), label)
_g(0, "male"); _g(1, "female"); _g(2, "child"); _g(3, "giant")

glots = OrderedDict()
def _t(_id, label): glots[str(_id)] = VoiceInfo(str(_id), label)
_t(0, "default"); _t(1, "male"); _t(2, "female"); _t(3, "child")
_t(4, "high"); _t(5, "mellow"); _t(6, "impulse"); _t(7, "odd"); _t(8, "colossus")

smodes = OrderedDict()
def _k(_id, label): smodes[str(_id)] = VoiceInfo(str(_id), label)
_k(0, "Natural"); _k(1, "Word-at-a-time"); _k(2, "Spelled")

class SynthDriver(SynthDriver):
    name = "sv"
    description = "SoftVoice (nvwave)"
    supportedSettings = (
        SynthDriver.RateSetting(), SynthDriver.VariantSetting(),
        SynthDriver.VoiceSetting(), SynthDriver.PitchSetting(),
        SynthDriver.InflectionSetting(),
        NumericDriverSetting("perturb", "Perturbation"),
        NumericDriverSetting("vfactor", "Vowel Factor"),
        NumericDriverSetting("avbias", "Voicing Gain"),
        NumericDriverSetting("afbias", "Frication Gain"),
        NumericDriverSetting("ahbias", "Aspiration Gain"),
        DriverSetting("intstyle", "Intonation Style"),
        DriverSetting("vmode", "Voicing Mode"),
        DriverSetting("gender", "Gender"),
        DriverSetting("glot", "Glottal Source"),
        DriverSetting("smode", "Speaking Mode"),
        NumericDriverSetting("pauseFactor", "Pause factor"),
    )
    supportedCommands = {IndexCommand}
    supportedNotifications = {synthIndexReached, synthDoneSpeaking}
    availableVoices = OrderedDict(
        (str(index + 1), VoiceInfo(str(index + 1), name))
        for index, name in enumerate(("English", "Spanish"))
    )

    @classmethod
    def check(cls):
        if ctypes.sizeof(ctypes.c_void_p) != 4: return False
        base_path = os.path.dirname(__file__)
        return bool(_find_tibase32(base_path)) and os.path.isfile(os.path.join(base_path, "softvoice_wrapper.dll"))

    def __init__(self):
        super().__init__()
        if ctypes.sizeof(ctypes.c_void_p) != 4:
            raise RuntimeError("SoftVoice: 32-bit only")

        self._acquireWrapper()

        try: self._outputDevice = config.conf["speech"]["outputDevice"]
        except: self._outputDevice = config.conf["audio"]["outputDevice"]

        self._player = None
        self._playerFormat = None
        self._bufSize = 65536
        self._audioBuf = ctypes.create_string_buffer(self._bufSize)
        self._tailHoldBytes = 0
        self._tailBuf = bytearray()
        
        self.speaking = False
        self._terminating = False

        self._bgQueue = queue.Queue()
        self._bgStop = threading.Event()
        self._bgThread = _BgThread(self._bgQueue, self._bgStop)
        self._bgThread.start()

        self._ratePercent = 50
        self._pitchPercent = 50
        self._inflectionPercent = 50
        self._perturbPercent = 0
        self._vfactorPercent = 20
        self._avbiasPercent = 50
        self._afbiasPercent = 50
        self._ahbiasPercent = 50
        self._pauseFactorPercent = 50  # Default 50
        
        self._variant = "0"
        self._intstyle = "0"
        self._vmode = "0"
        self._gender = "0"
        self._glot = "0"
        self._smode = "0"
        self.curvoice = "1"

        self._paramExplicit = {
            "intstyle": False, "vmode": False, "gender": False, "glot": False,
            "smode": False, "inflection": False, "perturb": False,
            "vfactor": False, "avbias": False, "afbias": False, "ahbias": False
        }

        self.rate = self._ratePercent
        self.pitch = 4
        self.inflection = 25
        self.perturb = self._perturbPercent
        self.vfactor = self._vfactorPercent
        self.avbias = self._avbiasPercent
        self.afbias = self._afbiasPercent
        self.ahbias = self._ahbiasPercent
        self.voice = self.curvoice
        self.pauseFactor = self._pauseFactorPercent

    def _acquireWrapper(self):
        global _wrapperHandle, _wrapperRefCount, _wrapperDll, _wrapperLock
        with _wrapperLock:
            self._base_path = os.path.dirname(__file__)
            self._tibase_path = _canonicalizeWindowsPath(_find_tibase32(self._base_path))
            self._wrapper_path = os.path.join(self._base_path, "softvoice_wrapper.dll")

            if not _wrapperHandle:
                if not _wrapperDll:
                    _wrapperDll = ctypes.cdll[self._wrapper_path]
                    self._setupPrototypes(_wrapperDll)
                _wrapperHandle = _wrapperDll.sv_initW(self._tibase_path, 1)
                if not _wrapperHandle: raise RuntimeError("Wrapper init failed")

            _wrapperRefCount += 1
            self._handle = _wrapperHandle
            self._dll = _wrapperDll
            self._hasPauseFactor = hasattr(self._dll, "sv_setPauseFactor")
            self._hasTrimSilence = hasattr(self._dll, "sv_setTrimSilence")

    def _setupPrototypes(self, dll):
        dll.sv_initW.argtypes = (ctypes.c_wchar_p, ctypes.c_int); dll.sv_initW.restype = ctypes.c_void_p
        dll.sv_free.argtypes = (ctypes.c_void_p,); dll.sv_stop.argtypes = (ctypes.c_void_p,)
        dll.sv_startSpeakW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p); dll.sv_startSpeakW.restype = ctypes.c_int
        dll.sv_read.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int), ctypes.c_void_p, ctypes.c_int)
        dll.sv_read.restype = ctypes.c_int
        for func in ["sv_setRate", "sv_setPitch", "sv_setF0Range", "sv_setF0Perturb", "sv_setVowelFactor", 
                     "sv_setAVBias", "sv_setAFBias", "sv_setAHBias", "sv_setPersonality", "sv_setF0Style", 
                     "sv_setVoicingMode", "sv_setGender", "sv_setGlottalSource", "sv_setSpeakingMode", "sv_setVoice"]:
            getattr(dll, func).argtypes = (ctypes.c_void_p, ctypes.c_int)
        dll.sv_getFormat.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
        dll.sv_getFormat.restype = ctypes.c_int
        if hasattr(dll, "sv_setPauseFactor"): dll.sv_setPauseFactor.argtypes = (ctypes.c_void_p, ctypes.c_int)
        if hasattr(dll, "sv_setTrimSilence"): dll.sv_setTrimSilence.argtypes = (ctypes.c_void_p, ctypes.c_int)

    def _enqueue(self, func, *args, **kwargs):
        if not self._terminating: self._bgQueue.put((func, args, kwargs))

    def terminate(self):
        global _wrapperRefCount, _wrapperHandle, _wrapperLock
        self._terminating = True
        self.cancel()
        try:
            self._bgStop.set()
            self._bgQueue.put(None)
            self._bgThread.join(timeout=2.0)
        except: pass

        with _wrapperLock:
            if _wrapperRefCount > 0: _wrapperRefCount -= 1
            if self._handle:
                try: self._dll.sv_stop(self._handle)
                except: pass
        self._handle = None
        self._dll = None
        self._player = None

    def cancel(self):
        self.speaking = False
        self._tailBuf = bytearray()
        if self._handle:
            try: self._dll.sv_stop(self._handle)
            except: pass
        if self._player: self._player.stop()
        try:
            while True:
                self._bgQueue.get_nowait()
                self._bgQueue.task_done()
        except queue.Empty: pass

    def pause(self, switch):
        if self._player: self._player.pause(switch)

    # --- Speaking ---
    def _buildBlocks(self, speechSequence):
        blocks = []
        textBuf = []
        def flush(indexesAfter):
            raw = " ".join(textBuf)
            textBuf.clear()
            safe = self._softVoiceSafeText(raw)
            blocks.append((safe, indexesAfter))
        for item in speechSequence:
            if isinstance(item, str): textBuf.append(item)
            elif isinstance(item, IndexCommand): flush([item.index])
        if textBuf: flush([])
        while blocks and (not blocks[-1][0]) and (not blocks[-1][1]): blocks.pop()
        anyText = any(bool(t) for (t, _) in blocks)
        allIndexes = []
        for (_, idxs) in blocks: allIndexes.extend(idxs)
        return blocks, anyText, allIndexes

    def _notifyIndexesAndDone(self, indexes):
        for i in indexes: synthIndexReached.notify(synth=self, index=i)
        synthDoneSpeaking.notify(synth=self)
        self.speaking = False

    def speak(self, speechSequence):
        if len(speechSequence) == 1 and isinstance(speechSequence[0], IndexCommand):
            self._enqueue(self._notifyIndexesAndDone, [speechSequence[0].index])
            return
        blocks, anyText, allIndexes = self._buildBlocks(speechSequence)
        if not anyText:
            self._enqueue(self._notifyIndexesAndDone, allIndexes)
            return
        self._enqueue(self._speakBg, blocks)

    def _tryCreatePlayerFromWrapper(self):
        if self._player: return True
        sr = ctypes.c_int(0); ch = ctypes.c_int(0); bits = ctypes.c_int(0)
        try:
            if self._dll.sv_getFormat(self._handle, ctypes.byref(sr), ctypes.byref(ch), ctypes.byref(bits)) != 1: raise Exception
            self._player = nvwave.WavePlayer(ch.value, sr.value, bits.value, outputDevice=self._outputDevice)
            self._playerFormat = (ch.value, sr.value, bits.value)
            self._recalcTailHoldBytes()
            return True
        except:
            try:
                self._player = nvwave.WavePlayer(1, 22050, 16, outputDevice=self._outputDevice)
                self._playerFormat = (1, 22050, 16)
                self._recalcTailHoldBytes()
                return True
            except: return False

    def _recalcTailHoldBytes(self):
        if not self._playerFormat: return
        ch, rate, bits = self._playerFormat
        bytesPerSec = int(rate * ch * (bits // 8))
        # Slider 100 = 160ms hold. Slider 0 = 0ms hold.
        holdMs = max(0, min(int(round(self._pauseFactorPercent * 1.6)), 200))
        holdBytes = int(bytesPerSec * holdMs / 1000.0)
        frameBytes = max(1, ch * (bits // 8))
        self._tailHoldBytes = (holdBytes // frameBytes) * frameBytes

    def _feedAudioWithTailHold(self, data: bytes):
        if not data or not self._player: return
        if self._tailHoldBytes <= 0:
            self._player.feed(data, len(data))
            return
        self._tailBuf.extend(data)
        if len(self._tailBuf) > self._tailHoldBytes:
            cut = len(self._tailBuf) - self._tailHoldBytes
            chunk = bytes(self._tailBuf[:cut])
            del self._tailBuf[:cut]
            if chunk: self._player.feed(chunk, len(chunk))

    def _flushTail(self):
        if self._player and self._tailBuf:
            chunk = bytes(self._tailBuf); self._tailBuf = bytearray()
            self._player.feed(chunk, len(chunk))

    def _pumpUntilDone(self):
        outType = ctypes.c_int(0); outValue = ctypes.c_int(0)
        playerReady = bool(self._player)
        
        while self.speaking:
            madeProgress = False
            while True:
                try: n = self._dll.sv_read(self._handle, ctypes.byref(outType), ctypes.byref(outValue), self._audioBuf, self._bufSize)
                except: return False
                t = outType.value
                if t == SV_ITEM_AUDIO:
                    if n > 0:
                        madeProgress = True
                        if not playerReady:
                            for _ in range(50):
                                if self._tryCreatePlayerFromWrapper(): playerReady = True; break
                                time.sleep(0.005)
                        if playerReady: self._feedAudioWithTailHold(self._audioBuf.raw[:n])
                    continue
                if t == SV_ITEM_DONE: return True
                if t == SV_ITEM_ERROR: return False
                break
            if not madeProgress: time.sleep(0.001)
        return False

    def _speakBg(self, blocks):
        self.speaking = True
        is_word_mode = (str(self._smode) == "1")
        for (text, indexesAfter) in blocks:
            if not self.speaking: break
            if text:
                text_segments = text.split(" ") if is_word_mode else [text[i : i + MAX_STRING_LENGTH] for i in range(0, len(text), MAX_STRING_LENGTH)]
                for seg in text_segments:
                    if not self.speaking: break
                    seg = seg.strip()
                    if not seg: continue
                    self._dll.sv_startSpeakW(self._handle, seg)
                    if not self._pumpUntilDone(): self.speaking = False; break
            if self.speaking:
                def cb(idxs=indexesAfter):
                    if self.speaking:
                        for i in idxs: synthIndexReached.notify(synth=self, index=i)
                if self._player: self._player.feed(b"", 0, onDone=cb); self._flushTail()
                else: cb()
        if not self.speaking: synthDoneSpeaking.notify(synth=self); return
        self._flushTail()
        def doneCb():
            if self.speaking: self.speaking = False; synthDoneSpeaking.notify(synth=self)
        if self._player: self._player.feed(b"", 0, onDone=doneCb); self._player.idle()

    def _softVoiceSafeText(self, s: str) -> str:
        s = _sanitizeText(s)
        if not s: return ""
        if self._pauseFactorPercent < 50:
            s = _labelColonRe.sub(r"\1 \2", s); s = _labelSemiRe.sub(r"\1 \2", s)
        if str(getattr(self, "_smode", "0")) == "2":
            def _spellMatch(m): return " ".join(list(m.group(0)))
            s = _spellWordRe.sub(_spellMatch, s)
        return " ".join(s.split()).strip()

    # --- Settings ---
    def _percentToParam(self, val, minVal, maxVal):
        ratio = float(val) / 100.0
        return int(round(minVal + (maxVal - minVal) * ratio))
    def _clampPercent(self, v): return max(0, min(100, int(v)))

    # Timbre Settings (protected)
    def _timbre_setter(self, name, func, minV, maxV, val):
        self._clampPercent(val)
        current = getattr(self, f"_{name}Percent")
        new_val = int(val)
        if self._variant != "0" and not self._paramExplicit[name]:
             setattr(self, f"_{name}Percent", new_val)
             if new_val != current:
                 self._paramExplicit[name] = True
                 if self._handle: getattr(self._dll, func)(self._handle, self._percentToParam(new_val, minV, maxV))
             return
        setattr(self, f"_{name}Percent", new_val)
        self._paramExplicit[name] = True
        if self._handle: getattr(self._dll, func)(self._handle, self._percentToParam(new_val, minV, maxV))

    def _get_rate(self): return int(self._ratePercent)
    def _set_rate(self, v):
        self._ratePercent = self._clampPercent(v)
        if self._handle: self._dll.sv_setRate(self._handle, self._percentToParam(self._ratePercent, 20, 500))

    def _get_pitch(self): return int(self._pitchPercent)
    def _set_pitch(self, v):
        self._pitchPercent = self._clampPercent(v)
        if self._handle: self._dll.sv_setPitch(self._handle, self._percentToParam(self._pitchPercent, 10, 2000))

    def _get_inflection(self): return int(self._inflectionPercent)
    def _set_inflection(self, v): self._timbre_setter("inflection", "sv_setF0Range", 0, 500, v)
    def _get_perturb(self): return int(self._perturbPercent)
    def _set_perturb(self, v): self._timbre_setter("perturb", "sv_setF0Perturb", 0, 500, v)
    def _get_vfactor(self): return int(self._vfactorPercent)
    def _set_vfactor(self, v): self._timbre_setter("vfactor", "sv_setVowelFactor", 0, 500, v)
    def _get_avbias(self): return int(self._avbiasPercent)
    def _set_avbias(self, v): self._timbre_setter("avbias", "sv_setAVBias", -50, 50, v)
    def _get_afbias(self): return int(self._afbiasPercent)
    def _set_afbias(self, v): self._timbre_setter("afbias", "sv_setAFBias", -50, 50, v)
    def _get_ahbias(self): return int(self._ahbiasPercent)
    def _set_ahbias(self, v): self._timbre_setter("ahbias", "sv_setAHBias", -50, 50, v)

    def _get_pauseFactor(self): return int(self._pauseFactorPercent)
    def _set_pauseFactor(self, v):
        self._pauseFactorPercent = self._clampPercent(v)
        if self._handle:
            if self._hasPauseFactor:
                # INVERSION FIX:
                # User 0 (None) -> Engine 100 (Fast/NoPause).
                # User 100 (Full) -> Engine 0 (Slow/Full).
                inverted = 100 - int(self._pauseFactorPercent)
                try: self._dll.sv_setPauseFactor(self._handle, inverted)
                except: pass
            if self._hasTrimSilence:
                # SECRET MODE: Enable trim only if slider is low (e.g. < 50)
                try: self._dll.sv_setTrimSilence(self._handle, 1 if self._pauseFactorPercent < 50 else 0)
                except: pass
        self._recalcTailHoldBytes()

    def _get_availableVariants(self): return variants
    def _get_variant(self): return getattr(self, "_variant", "0")
    def _set_variant(self, _id):
        new_v = str(_id)
        if new_v != self._variant:
            for k in self._paramExplicit: self._paramExplicit[k] = False
        self._variant = new_v
        if self._handle: self._dll.sv_setPersonality(self._handle, int(_id))

    def _get_voice(self): return getattr(self, "curvoice", "1")
    def _set_voice(self, v):
        self.curvoice = str(v)
        if self._handle: self._dll.sv_setVoice(self._handle, int(v))

    def _set_enum_generic(self, attr_name, func_name, val_id):
        key = attr_name.strip("_")
        val = int(val_id)
        current = int(getattr(self, attr_name, 0))
        setattr(self, attr_name, str(val_id))
        if self._variant != "0" and not self._paramExplicit.get(key, False):
            if val != current:
                self._paramExplicit[key] = True
                if self._handle: getattr(self._dll, func_name)(self._handle, val)
            return
        self._paramExplicit[key] = True
        if self._handle: getattr(self._dll, func_name)(self._handle, val)

    def _get_availableIntstyles(self): return intstyles
    def _get_intstyle(self): return getattr(self, "_intstyle", "0")
    def _set_intstyle(self, v): self._set_enum_generic("_intstyle", "sv_setF0Style", v)
    def _get_availableVmodes(self): return vmodes
    def _get_vmode(self): return getattr(self, "_vmode", "0")
    def _set_vmode(self, v): self._set_enum_generic("_vmode", "sv_setVoicingMode", v)
    def _get_availableGenders(self): return genders
    def _get_gender(self): return getattr(self, "_gender", "0")
    def _set_gender(self, v): self._set_enum_generic("_gender", "sv_setGender", v)
    def _get_availableGlots(self): return glots
    def _get_glot(self): return getattr(self, "_glot", "0")
    def _set_glot(self, v): self._set_enum_generic("_glot", "sv_setGlottalSource", v)
    def _get_availableSmodes(self): return smodes
    def _get_smode(self): return getattr(self, "_smode", "0")
    def _set_smode(self, v): self._set_enum_generic("_smode", "sv_setSpeakingMode", v)
