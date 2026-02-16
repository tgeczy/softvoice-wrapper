"""Client-side helper for SoftVoice speech synthesis.

On 32-bit Python (NVDA 2025.3 and earlier): loads softvoice_wrapper.dll
directly via ctypes -- no host process needed.

On 64-bit Python (NVDA 2026.1+): communicates with a 32-bit host process
over IPC so the 32-bit wrapper DLL can be used from 64-bit NVDA.

Adapted from Fastfinge's Eloquence 64 project with permission.
Original: https://github.com/Fastfinge/eloquence_64
"""
from __future__ import annotations

import ctypes
import itertools
import logging
import os
import queue
import subprocess
import threading
import time
from typing import Any, Callable, Dict, Optional, Tuple

IS_64BIT = ctypes.sizeof(ctypes.c_void_p) == 8

if IS_64BIT:
	from . import _ipc

LOGGER = logging.getLogger(__name__)

# Stream item types from softvoice_wrapper.h
SV_ITEM_NONE = 0
SV_ITEM_AUDIO = 1
SV_ITEM_DONE = 2
SV_ITEM_ERROR = 3

HOST_EXECUTABLE = "softvoice_host32.exe"
HOST_SCRIPT = "_host_softvoice32.py"
AUTH_KEY_BYTES = 16

# DLL setter functions allowed for dll_call
_ALLOWED_DLL_CALLS = frozenset({
	"sv_setRate", "sv_setPitch", "sv_setF0Range", "sv_setF0Perturb",
	"sv_setVowelFactor", "sv_setAVBias", "sv_setAFBias", "sv_setAHBias",
	"sv_setPersonality", "sv_setF0Style", "sv_setVoicingMode", "sv_setGender",
	"sv_setGlottalSource", "sv_setSpeakingMode", "sv_setVoice",
	"sv_setPauseFactor", "sv_setTrimSilence", "sv_setMaxLeadMs",
})

AudioChunk = Tuple[bytes, Optional[int], bool, int]  # (data, index, is_final, seq)


def _convert_audio(data: bytes) -> bytes:
	"""Convert 8-bit unsigned PCM to 16-bit signed PCM.

	The SoftVoice engine outputs 8-bit unsigned PCM.
	NVDA's WASAPI backend can't handle 8-bit, so we convert to 16-bit.
	WASAPI handles the 11025 Hz resampling natively.
	"""
	import array
	return array.array('h', ((b - 128) << 8 for b in data)).tobytes()


# ---------------------------------------------------------------------------
# Audio handling (shared by both 32-bit and 64-bit modes)
# ---------------------------------------------------------------------------

class AudioWorker(threading.Thread):
	"""Pulls audio events from the queue and feeds them to nvwave.WavePlayer."""

	def __init__(self, player, audio_queue: "queue.Queue[Optional[AudioChunk]]",
				 get_sequence: Callable[[], int], convert_8to16: bool = False,
				 player_lock: Optional[threading.RLock] = None,
				 auto_idle: bool = True):
		super().__init__(daemon=True, name="SoftVoiceAudioWorker")
		self._player = player
		self._queue = audio_queue
		self._get_sequence = get_sequence
		self._convert_8to16 = convert_8to16
		self._running = True
		self._stopping = False
		self._player_lock = player_lock or threading.RLock()
		self._auto_idle = auto_idle

	def run(self) -> None:
		while self._running:
			try:
				chunk = self._queue.get(timeout=0.1)
			except queue.Empty:
				continue
			if chunk is None:
				break

			data, index, is_final, seq = chunk

			if seq < self._get_sequence():
				self._queue.task_done()
				continue

			if not data and index is None:
				if is_final and self._auto_idle:
					with self._player_lock:
						if not self._stopping:
							self._player.idle()
					if not self._stopping:
						self._invoke_done_callback()
				self._queue.task_done()
				continue

			if self._stopping:
				self._queue.task_done()
				continue

			try:
				if self._convert_8to16 and data:
					data = _convert_audio(data)
				with self._player_lock:
					if not self._stopping and self._player:
						self._player.feed(data)
			except FileNotFoundError:
				LOGGER.warning("Sound device not found during feed")
			except Exception:
				LOGGER.exception("WavePlayer feed failed")
			self._queue.task_done()

	def stop(self) -> None:
		self._stopping = True
		self._running = False
		self._queue.put(None)

	def _invoke_done_callback(self) -> None:
		if _on_done:
			try:
				_on_done()
			except Exception:
				LOGGER.exception("Done callback failed")


# ---------------------------------------------------------------------------
# 32-bit direct client (loads DLL in-process)
# ---------------------------------------------------------------------------

class SoftVoiceDirectClient:
	"""Direct ctypes access to softvoice_wrapper.dll for 32-bit NVDA."""

	def __init__(self) -> None:
		self._dll = None
		self._handle = None
		self._audio_queue: "queue.Queue[Optional[AudioChunk]]" = queue.Queue()
		self._player = None
		self._player_lock = threading.RLock()
		self._audio_worker: Optional[AudioWorker] = None
		self._should_stop = False
		self._sequence = 0
		self._current_seq = 0
		# Audio format
		self._sample_rate = 0
		self._channels = 0
		self._bits_per_sample = 0
		# Feature flags (cached after first init)
		self._has_pause_factor = False
		self._has_trim_silence = False
		# Read buffer
		self._buf_size = 65536
		self._audio_buf = None
		self._out_type = None
		self._out_value = None

	def do_initialize(self, dll_path: str, tibase_path: str, initial_voice: int) -> Dict[str, Any]:
		"""Load the wrapper DLL and initialize the engine."""
		if self._dll is None:
			self._dll = ctypes.cdll.LoadLibrary(dll_path)
			self._setup_ctypes()

		self._audio_buf = ctypes.create_string_buffer(self._buf_size)
		self._out_type = ctypes.c_int(0)
		self._out_value = ctypes.c_int(0)

		self._handle = self._dll.sv_initW(tibase_path, initial_voice)
		if not self._handle:
			raise RuntimeError("sv_initW returned NULL")

		self._has_pause_factor = hasattr(self._dll, "sv_setPauseFactor")
		self._has_trim_silence = hasattr(self._dll, "sv_setTrimSilence")

		# Query audio format
		sr = ctypes.c_int(0)
		ch = ctypes.c_int(0)
		bps = ctypes.c_int(0)
		if self._dll.sv_getFormat(self._handle, ctypes.byref(sr), ctypes.byref(ch), ctypes.byref(bps)):
			self._sample_rate = sr.value
			self._channels = ch.value
			self._bits_per_sample = bps.value
		else:
			self._sample_rate = 22050
			self._channels = 1
			self._bits_per_sample = 16

		return {
			"format": {
				"sampleRate": self._sample_rate,
				"channels": self._channels,
				"bitsPerSample": self._bits_per_sample,
			},
			"hasPauseFactor": self._has_pause_factor,
			"hasTrimSilence": self._has_trim_silence,
		}

	def _setup_ctypes(self):
		dll = self._dll
		dll.sv_initW.argtypes = (ctypes.c_wchar_p, ctypes.c_int)
		dll.sv_initW.restype = ctypes.c_void_p
		dll.sv_free.argtypes = (ctypes.c_void_p,)
		dll.sv_free.restype = None
		dll.sv_stop.argtypes = (ctypes.c_void_p,)
		dll.sv_stop.restype = None
		dll.sv_startSpeakW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p)
		dll.sv_startSpeakW.restype = ctypes.c_int
		dll.sv_read.argtypes = (
			ctypes.c_void_p,
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
			ctypes.c_void_p,
			ctypes.c_int,
		)
		dll.sv_read.restype = ctypes.c_int
		dll.sv_getFormat.argtypes = (
			ctypes.c_void_p,
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
		)
		dll.sv_getFormat.restype = ctypes.c_int
		for func_name in _ALLOWED_DLL_CALLS:
			if hasattr(dll, func_name):
				fn = getattr(dll, func_name)
				fn.argtypes = (ctypes.c_void_p, ctypes.c_int)
				fn.restype = None

	# ------------------------------------------------------------------
	# Audio
	def initialize_audio(self, channels: int, sample_rate: int, bits_per_sample: int) -> None:
		if self._player:
			return
		import nvwave
		import config
		try:
			from buildVersion import version_year
		except ImportError:
			version_year = 2025

		# SoftVoice engine outputs 8-bit unsigned PCM at 11025 Hz.
		# WASAPI can't handle 8-bit, so convert to 16-bit.
		# Let WASAPI handle the 11025 Hz resampling natively.
		convert_8to16 = (bits_per_sample == 8)
		player_bps = 16 if convert_8to16 else bits_per_sample

		if version_year >= 2025:
			device = config.conf["audio"]["outputDevice"]
			player = nvwave.WavePlayer(channels, sample_rate, player_bps,
									   outputDevice=device)
		else:
			device = config.conf["speech"]["outputDevice"]
			player = nvwave.WavePlayer(channels, sample_rate, player_bps,
									   outputDevice=device, buffered=True)
		self._player = player
		self._audio_worker = AudioWorker(player, self._audio_queue,
										 lambda: self._sequence,
										 convert_8to16=convert_8to16,
										 player_lock=self._player_lock,
										 auto_idle=False)
		self._audio_worker.start()

	# ------------------------------------------------------------------
	# Speech (blocking)
	def do_speak(self, text: str) -> bool:
		"""Start speech and pump read loop. Returns True on success."""
		self._should_stop = False
		self._current_seq = self._sequence
		rc = self._dll.sv_startSpeakW(self._handle, text)
		if rc != 0:
			LOGGER.error("sv_startSpeakW returned %d", rc)
			self._audio_queue.put((b"", None, True, self._current_seq))
			return False
		return self._read_loop()

	def _read_loop(self) -> bool:
		"""Poll sv_read() and push audio to queue. Returns True if completed normally."""
		while not self._should_stop:
			try:
				n = self._dll.sv_read(
					self._handle,
					ctypes.byref(self._out_type),
					ctypes.byref(self._out_value),
					self._audio_buf,
					self._buf_size,
				)
			except Exception:
				LOGGER.exception("sv_read crashed")
				self._audio_queue.put((b"", None, True, self._current_seq))
				return False

			t = self._out_type.value

			if t == SV_ITEM_AUDIO and n > 0:
				self._audio_queue.put((bytes(self._audio_buf.raw[:n]), None, False, self._current_seq))
			elif t == SV_ITEM_DONE:
				self._audio_queue.put((b"", None, True, self._current_seq))
				return True
			elif t == SV_ITEM_ERROR:
				LOGGER.error("Wrapper error %d", self._out_value.value)
				self._audio_queue.put((b"", None, True, self._current_seq))
				return False
			elif t == SV_ITEM_NONE:
				time.sleep(0.001)
		return False

	# ------------------------------------------------------------------
	# Control
	def dll_call(self, func_name: str, value: int) -> None:
		if func_name not in _ALLOWED_DLL_CALLS:
			raise ValueError(f"Disallowed: {func_name}")
		fn = getattr(self._dll, func_name, None)
		if fn:
			fn(self._handle, value)

	def stop(self) -> None:
		self._sequence += 1
		self._should_stop = True
		if self._audio_worker:
			self._audio_worker._stopping = True
		if self._handle:
			self._dll.sv_stop(self._handle)
		# Call player.stop() WITHOUT the lock.  WavePlayer.stop() is
		# thread-safe (NVDA calls it from MainThread while synths feed
		# from background threads).  Using the lock here would deadlock
		# if AudioWorker is blocked inside feed() on WASAPI buffer space.
		if self._player:
			try:
				self._player.stop()
			except Exception:
				LOGGER.exception("WavePlayer stop failed")

	def pause(self, switch: bool) -> None:
		with self._player_lock:
			if self._player:
				self._player.pause(switch)

	def feed_marker(self, on_done=None) -> None:
		with self._player_lock:
			if self._player:
				self._player.feed(b"", onDone=on_done)

	def player_idle(self) -> None:
		with self._player_lock:
			if self._player:
				self._player.idle()

	def get_format(self) -> Dict[str, int]:
		return {
			"sampleRate": self._sample_rate,
			"channels": self._channels,
			"bitsPerSample": self._bits_per_sample,
		}

	# ------------------------------------------------------------------
	# Shutdown
	def shutdown(self) -> None:
		if self._audio_worker:
			self._audio_worker.stop()
			self._audio_worker.join(timeout=1)
			self._audio_worker = None
		with self._player_lock:
			if self._player:
				self._player.close()
				self._player = None
		if self._handle and self._dll:
			try:
				self._dll.sv_free(self._handle)
			except Exception:
				LOGGER.exception("sv_free failed")
			self._handle = None
		# Reset audio state for re-initialization.
		# Keep self._dll alive — the wrapper DLL stays loaded so that
		# MinHook trampolines remain valid, but sv_free's force-unload
		# ensures the engine DLLs (tibase32 etc.) are fully released.
		self._audio_queue = queue.Queue()
		self._sequence = 0
		self._current_seq = 0


# ---------------------------------------------------------------------------
# 64-bit IPC client (communicates with 32-bit host process)
# ---------------------------------------------------------------------------

if IS_64BIT:
	from dataclasses import dataclass

	@dataclass
	class HostProcess:
		process: subprocess.Popen
		connection: _ipc.IpcConnection
		listener: Any  # socket

	class SoftVoiceHostClient:
		"""Spawns a 32-bit host process and communicates via IPC."""

		def __init__(self) -> None:
			self._host: Optional[HostProcess] = None
			self._pending: Dict[int, threading.Event] = {}
			self._responses: Dict[int, Dict[str, Any]] = {}
			self._receiver: Optional[threading.Thread] = None
			self._id_counter = itertools.count(1)
			self._audio_queue: "queue.Queue[Optional[AudioChunk]]" = queue.Queue()
			self._player = None
			self._player_lock = threading.RLock()
			self._audio_worker: Optional[AudioWorker] = None
			self._send_lock = threading.Lock()
			self._sequence = 0
			self._current_seq = 0

		def ensure_started(self) -> None:
			if self._host:
				return
			addon_dir = os.path.abspath(os.path.dirname(__file__))
			authkey = os.urandom(AUTH_KEY_BYTES)

			listener = _ipc.create_listener()
			port = listener.getsockname()[1]

			cmd = list(self._resolve_host_executable(addon_dir))
			cmd.extend([
				"--address", f"127.0.0.1:{port}",
				"--authkey", authkey.hex(),
				"--log-dir", addon_dir,
			])
			LOGGER.info("Launching SoftVoice host: %s", cmd)
			proc = subprocess.Popen(cmd, cwd=addon_dir, creationflags=subprocess.CREATE_NO_WINDOW)

			conn = _ipc.accept_authenticated(listener, authkey)
			self._host = HostProcess(process=proc, connection=conn, listener=listener)

			self._receiver = threading.Thread(target=self._receiver_loop, daemon=True,
											  name="SoftVoiceReceiver")
			self._receiver.start()

		def _resolve_host_executable(self, addon_dir: str):
			override = os.environ.get("SOFTVOICE_HOST_COMMAND")
			if override:
				import shlex
				return shlex.split(override)
			exe_path = os.path.join(addon_dir, HOST_EXECUTABLE)
			if os.path.exists(exe_path):
				return [exe_path]
			script_path = os.path.join(addon_dir, HOST_SCRIPT)
			if os.path.exists(script_path):
				return ["py", "-3.14-32", script_path]
			raise RuntimeError("SoftVoice host executable not found")

		def initialize_audio(self, channels: int, sample_rate: int, bits_per_sample: int) -> None:
			if self._player:
				return
			import nvwave
			import config
			try:
				from buildVersion import version_year
			except ImportError:
				version_year = 2025

			# SoftVoice engine outputs 8-bit unsigned PCM at 11025 Hz.
			# WASAPI can't handle 8-bit, so convert to 16-bit.
			# Let WASAPI handle the 11025 Hz resampling natively.
			convert_8to16 = (bits_per_sample == 8)
			player_bps = 16 if convert_8to16 else bits_per_sample

			if version_year >= 2025:
				device = config.conf["audio"]["outputDevice"]
				player = nvwave.WavePlayer(channels, sample_rate, player_bps,
										   outputDevice=device)
			else:
				device = config.conf["speech"]["outputDevice"]
				player = nvwave.WavePlayer(channels, sample_rate, player_bps,
										   outputDevice=device, buffered=True)
			self._player = player
			self._audio_worker = AudioWorker(player, self._audio_queue,
											 lambda: self._sequence,
											 convert_8to16=convert_8to16,
											 player_lock=self._player_lock,
											 auto_idle=False)
			self._audio_worker.start()

		def _receiver_loop(self) -> None:
			connection = self._host.connection if self._host else None
			if connection is None:
				return
			while True:
				try:
					message = connection.recv()
				except (EOFError, ConnectionAbortedError, OSError):
					LOGGER.info("Host connection closed")
					for msg_id, event in list(self._pending.items()):
						self._responses[msg_id] = {"error": "connectionClosed"}
						event.set()
					self._pending.clear()
					break
				except Exception:
					LOGGER.exception("Unexpected error in receiver loop")
					for msg_id, event in list(self._pending.items()):
						self._responses[msg_id] = {"error": "receiverException"}
						event.set()
					self._pending.clear()
					break

				msg_type = message.get("type")
				if msg_type == "response":
					msg_id = message["id"]
					self._responses[msg_id] = message
					event = self._pending.pop(msg_id, None)
					if event:
						event.set()
				elif msg_type == "event":
					self._handle_event(message["event"], message.get("payload", {}))
				else:
					LOGGER.warning("Unknown message type %s", msg_type)

		def _handle_event(self, event: str, payload: Dict[str, Any]) -> None:
			if event == "audio":
				data = payload.get("data", b"")
				is_final = bool(payload.get("final", False))
				seq = self._current_seq
				self._audio_queue.put((data, None, is_final, seq))
			elif event == "stopped":
				LOGGER.debug("Host reported stopped event")

		def send_command(self, command: str, timeout: float = 10.0, **payload: Any) -> Dict[str, Any]:
			if not self._host:
				raise RuntimeError("Host not started")
			msg_id = next(self._id_counter)
			event = threading.Event()
			self._pending[msg_id] = event
			with self._send_lock:
				try:
					self._host.connection.send({
						"type": "command", "id": msg_id,
						"command": command, "payload": payload,
					})
				except Exception:
					self._pending.pop(msg_id, None)
					raise
			if not event.wait(timeout=timeout):
				self._pending.pop(msg_id, None)
				LOGGER.error("Command %s timed out after %.1f seconds", command, timeout)
				raise RuntimeError(f"Command {command} timed out")
			response = self._responses.pop(msg_id, {"error": "no response received"})
			if "error" in response:
				raise RuntimeError(response["error"])
			return response.get("payload", {})

		def stop(self) -> None:
			if not self._host:
				return
			self._sequence += 1
			# 1. Signal AudioWorker to stop feeding (prevents new feed calls)
			if self._audio_worker:
				self._audio_worker._stopping = True
			# 2. Immediately silence the player WITHOUT the lock.
			#    WavePlayer.stop() is thread-safe (NVDA calls it from
			#    MainThread while synths feed from background threads).
			#    Using the lock would deadlock if AudioWorker is blocked
			#    inside feed() waiting for WASAPI buffer space.
			if self._player:
				try:
					self._player.stop()
				except Exception:
					LOGGER.exception("WavePlayer stop failed")
			# 3. Drain audio queue so AudioWorker doesn't feed stale data
			while not self._audio_queue.empty():
				try:
					self._audio_queue.get_nowait()
					self._audio_queue.task_done()
				except queue.Empty:
					break
			# 4. Fire-and-forget stop command to the host (don't block)
			try:
				msg_id = next(self._id_counter)
				with self._send_lock:
					self._host.connection.send({
						"type": "command", "id": msg_id,
						"command": "stop", "payload": {},
					})
			except Exception:
				LOGGER.exception("Stop command failed")

		def pause(self, switch: bool) -> None:
			with self._player_lock:
				if self._player:
					self._player.pause(switch)

		def feed_marker(self, on_done=None) -> None:
			with self._player_lock:
				if self._player:
					self._player.feed(b"", onDone=on_done)

		def player_idle(self) -> None:
			with self._player_lock:
				if self._player:
					self._player.idle()

		def shutdown(self) -> None:
			if not self._host:
				return
			if self._audio_worker:
				self._audio_worker.stop()
				self._audio_worker.join(timeout=1)
				self._audio_worker = None
			with self._player_lock:
				if self._player:
					self._player.close()
					self._player = None
			try:
				self.send_command("delete", timeout=3.0)
			except Exception:
				LOGGER.exception("Failed to delete host cleanly")
			if self._receiver:
				self._receiver.join(timeout=2)
				self._receiver = None
			try:
				self._host.connection.close()
			except Exception:
				pass
			try:
				self._host.listener.close()
			except Exception:
				pass
			try:
				self._host.process.terminate()
				self._host.process.wait(timeout=2)
			except Exception:
				LOGGER.exception("Failed to terminate host process")
				try:
					self._host.process.kill()
				except Exception:
					pass
			self._host = None
			# Reset state so ensure_started() / initialize can be called again
			self._audio_queue = queue.Queue()
			self._sequence = 0
			self._current_seq = 0
			self._pending.clear()
			self._responses.clear()
			self._id_counter = itertools.count(1)


# ---------------------------------------------------------------------------
# Module-level singleton and public API
# ---------------------------------------------------------------------------

_client: Any = SoftVoiceHostClient() if IS_64BIT else SoftVoiceDirectClient()
_on_done: Optional[Callable] = None
_format: Dict[str, int] = {}
_has_pause_factor: bool = False
_has_trim_silence: bool = False


def initialize(done_callback=None) -> Dict[str, Any]:
	"""Start the host process (64-bit) or load the DLL directly (32-bit)."""
	global _on_done, _format, _has_pause_factor, _has_trim_silence
	_on_done = done_callback

	addon_dir = os.path.abspath(os.path.dirname(__file__))
	tibase_path = _find_tibase32(addon_dir)
	dll_path = os.path.join(addon_dir, "softvoice_wrapper.dll")

	if IS_64BIT:
		_client.ensure_started()
		result = _client.send_command(
			"initialize",
			dllPath=dll_path,
			tibasePath=tibase_path,
			initialVoice=1,
		)
	else:
		result = _client.do_initialize(dll_path, tibase_path, 1)

	_format = result.get("format", {})
	_has_pause_factor = result.get("hasPauseFactor", False)
	_has_trim_silence = result.get("hasTrimSilence", False)

	_client.initialize_audio(
		channels=_format.get("channels", 1),
		sample_rate=_format.get("sampleRate", 22050),
		bits_per_sample=_format.get("bitsPerSample", 16),
	)

	return result


def has_pause_factor() -> bool:
	return _has_pause_factor


def has_trim_silence() -> bool:
	return _has_trim_silence


def speak(text: str) -> bool:
	"""Single-chunk speech (blocks until done). Returns True on success."""
	# Reset AudioWorker's stopping flag so it will feed new audio
	if _client._audio_worker:
		_client._audio_worker._stopping = False
	if IS_64BIT:
		_client._current_seq = _client._sequence
		try:
			_client.send_command("speak", text=text, timeout=30.0)
			return True
		except Exception:
			LOGGER.exception("speak command failed")
			return False
	else:
		return _client.do_speak(text)


def dll_call(func_name: str, value: int) -> None:
	"""Call a sv_set* function by name (parameter setters)."""
	if IS_64BIT:
		try:
			_client.send_command("dllCall", funcName=func_name, value=value)
		except Exception:
			LOGGER.exception("dllCall %s failed", func_name)
	else:
		_client.dll_call(func_name, value)


def stop() -> None:
	_client.stop()


def pause(switch: bool) -> None:
	_client.pause(switch)


def feed_marker(on_done=None) -> None:
	"""Feed an empty marker to the player. on_done fires when it reaches playback."""
	_client.feed_marker(on_done)


def player_idle() -> None:
	"""Signal the player that no more audio is coming."""
	_client.player_idle()


def terminate() -> None:
	_client.shutdown()


def get_format() -> Dict[str, int]:
	return dict(_format)


def check() -> bool:
	"""Check if the required files are available."""
	base = os.path.abspath(os.path.dirname(__file__))
	wrapper_dll = os.path.join(base, "softvoice_wrapper.dll")
	tibase_dll = _find_tibase32(base)

	if not os.path.isfile(wrapper_dll) or not tibase_dll:
		return False

	if IS_64BIT:
		host_exe = os.path.join(base, HOST_EXECUTABLE)
		host_script = os.path.join(base, HOST_SCRIPT)
		return os.path.exists(host_exe) or os.path.exists(host_script)
	else:
		return True


def _find_tibase32(base_path: str) -> str:
	for p in (
		os.path.join(base_path, "tibase32.dll"),
		os.path.join(base_path, "TIBASE32.DLL"),
	):
		if os.path.isfile(p):
			return p
	return ""
