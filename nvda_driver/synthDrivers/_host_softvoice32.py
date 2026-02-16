"""32-bit host process for SoftVoice speech synthesis.

This module runs as a separate 32-bit Python process.  It loads the
softvoice_wrapper.dll (which in turn loads the 32-bit tibase32.dll)
and exposes an RPC protocol over a TCP socket so that 64-bit NVDA can
use the synthesizer.

Adapted from Fastfinge's Eloquence 64 project with permission.
Original: https://github.com/Fastfinge/eloquence_64
"""
from __future__ import annotations

import argparse
import ctypes
import logging
import os
import threading
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional

import _ipc

# Stream item types from softvoice_wrapper.h
SV_ITEM_NONE = 0
SV_ITEM_AUDIO = 1
SV_ITEM_DONE = 2
SV_ITEM_ERROR = 3

LOGGER = logging.getLogger("softvoice.host")

# DLL setter functions that the client is allowed to call via dllCall
_ALLOWED_DLL_CALLS = frozenset({
	"sv_setRate", "sv_setPitch", "sv_setF0Range", "sv_setF0Perturb",
	"sv_setVowelFactor", "sv_setAVBias", "sv_setAFBias", "sv_setAHBias",
	"sv_setPersonality", "sv_setF0Style", "sv_setVoicingMode", "sv_setGender",
	"sv_setGlottalSource", "sv_setSpeakingMode", "sv_setVoice",
	"sv_setPauseFactor", "sv_setTrimSilence", "sv_setMaxLeadMs",
})


def configure_logging(log_dir: Optional[str]) -> None:
	logging.basicConfig(
		filename=os.path.join(log_dir, "softvoice-host.log") if log_dir else None,
		level=logging.DEBUG,
		format="%(asctime)s %(levelname)s %(message)s",
	)


@dataclass
class HostConfig:
	dll_path: str          # Path to softvoice_wrapper.dll
	tibase_path: str       # Path to tibase32.dll (passed to sv_initW)
	initial_voice: int     # Initial voice selection (1=English, 2=Spanish)


class SoftVoiceRuntime:
	"""Wraps access to the 32-bit softvoice_wrapper.dll."""

	def __init__(self, conn: _ipc.IpcConnection, config: HostConfig):
		self._conn = conn
		self._config = config
		self._dll = None
		self._handle = None
		self._should_stop = False
		# Audio format (filled after init)
		self._sample_rate = 0
		self._channels = 0
		self._bits_per_sample = 0
		# Read buffer
		self._buf_size = 65536
		self._audio_buf = ctypes.create_string_buffer(self._buf_size)
		self._out_type = ctypes.c_int(0)
		self._out_value = ctypes.c_int(0)

	def _send_event(self, event: str, **payload: object) -> None:
		try:
			self._conn.send({"type": "event", "event": event, "payload": payload})
		except Exception:
			LOGGER.exception("Failed to send event %s", event)

	# ------------------------------------------------------------------
	# Initialization
	def start(self) -> Dict[str, Any]:
		LOGGER.info("Loading softvoice_wrapper.dll from %s", self._config.dll_path)
		self._dll = ctypes.cdll.LoadLibrary(self._config.dll_path)
		self._setup_ctypes()

		LOGGER.info("Calling sv_initW with tibase_path=%s, voice=%d",
					self._config.tibase_path, self._config.initial_voice)
		self._handle = self._dll.sv_initW(self._config.tibase_path, self._config.initial_voice)
		if not self._handle:
			raise RuntimeError("sv_initW returned NULL")

		# Detect optional features
		has_pause_factor = hasattr(self._dll, "sv_setPauseFactor")
		has_trim_silence = hasattr(self._dll, "sv_setTrimSilence")

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
		LOGGER.info("Audio format: %d Hz, %d ch, %d bps",
					self._sample_rate, self._channels, self._bits_per_sample)

		return {
			"format": {
				"sampleRate": self._sample_rate,
				"channels": self._channels,
				"bitsPerSample": self._bits_per_sample,
			},
			"hasPauseFactor": has_pause_factor,
			"hasTrimSilence": has_trim_silence,
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
		# All setter functions: (handle, int) -> void
		for func_name in _ALLOWED_DLL_CALLS:
			if hasattr(dll, func_name):
				fn = getattr(dll, func_name)
				fn.argtypes = (ctypes.c_void_p, ctypes.c_int)
				fn.restype = None

	# ------------------------------------------------------------------
	# Speech
	def speak(self, text: str) -> None:
		"""Start speech and pump audio until done."""
		self._should_stop = False
		rc = self._dll.sv_startSpeakW(self._handle, text)
		if rc != 0:
			LOGGER.error("sv_startSpeakW returned %d", rc)
			self._send_event("audio", data=b"", index=None, final=True)
			return
		self._read_loop()

	def _read_loop(self) -> None:
		"""Pull items from the wrapper and push them as events to the client."""
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
				self._send_event("audio", data=b"", index=None, final=True)
				return

			t = self._out_type.value

			if t == SV_ITEM_AUDIO and n > 0:
				self._send_event("audio", data=self._audio_buf.raw[:n], index=None, final=False)
			elif t == SV_ITEM_DONE:
				self._send_event("audio", data=b"", index=None, final=True)
				return
			elif t == SV_ITEM_ERROR:
				LOGGER.error("Wrapper error %d", self._out_value.value)
				self._send_event("audio", data=b"", index=None, final=True)
				return
			elif t == SV_ITEM_NONE:
				time.sleep(0.001)

	# ------------------------------------------------------------------
	# Control
	def stop(self) -> None:
		self._should_stop = True
		if self._handle:
			self._dll.sv_stop(self._handle)
		self._send_event("stopped")

	def dll_call(self, func_name: str, value: int) -> None:
		"""Call a whitelisted sv_set* function."""
		if func_name not in _ALLOWED_DLL_CALLS:
			raise ValueError(f"Disallowed DLL call: {func_name}")
		fn = getattr(self._dll, func_name, None)
		if fn is None:
			raise AttributeError(f"DLL has no function: {func_name}")
		fn(self._handle, value)

	def get_format(self) -> Dict[str, int]:
		return {
			"sampleRate": self._sample_rate,
			"channels": self._channels,
			"bitsPerSample": self._bits_per_sample,
		}

	def delete(self) -> None:
		if self._handle:
			LOGGER.info("Freeing SoftVoice handle")
			self._dll.sv_free(self._handle)
			self._handle = None


class HostController:
	"""Receives commands from the 64-bit NVDA client and dispatches them.

	Speech commands run in a worker thread so that the main recv loop
	remains responsive to stop commands.
	"""

	def __init__(self, conn: _ipc.IpcConnection):
		self._conn = conn
		self._runtime: Optional[SoftVoiceRuntime] = None
		self._should_exit = False
		self._speak_thread: Optional[threading.Thread] = None
		self._handlers = {
			"initialize": self._handle_initialize,
			"speak": self._handle_speak,
			"stop": self._handle_stop,
			"dllCall": self._handle_dll_call,
			"getFormat": self._handle_get_format,
			"delete": self._handle_delete,
		}

	def serve_forever(self) -> None:
		LOGGER.info("Host controller waiting for commands")
		while not self._should_exit:
			try:
				message = self._conn.recv()
			except (EOFError, ConnectionError, OSError) as exc:
				LOGGER.info("Connection closed: %s", exc)
				break
			if not isinstance(message, dict):
				LOGGER.warning("Unexpected message %r", message)
				continue
			msg_type = message.get("type")
			if msg_type != "command":
				LOGGER.warning("Unsupported message type %s", msg_type)
				continue
			msg_id = message.get("id")
			command = message.get("command")
			handler = self._handlers.get(command)
			if handler is None:
				LOGGER.error("Unknown command %s", command)
				self._conn.send({"type": "response", "id": msg_id, "error": "unknownCommand"})
				continue

			if command == "speak":
				self._wait_for_speak_thread()
				self._speak_thread = threading.Thread(
					target=self._run_blocking_handler,
					args=(msg_id, handler, message.get("payload", {})),
					daemon=True,
				)
				self._speak_thread.start()
			else:
				try:
					payload = handler(**message.get("payload", {}))
					self._conn.send({"type": "response", "id": msg_id, "payload": payload or {}})
					if command == "delete" and self._should_exit:
						break
				except Exception as exc:
					LOGGER.exception("Command %s failed", command)
					self._conn.send({"type": "response", "id": msg_id, "error": str(exc)})

	def _run_blocking_handler(self, msg_id: int, handler, payload: Dict[str, Any]) -> None:
		try:
			result = handler(**payload)
			self._conn.send({"type": "response", "id": msg_id, "payload": result or {}})
		except Exception as exc:
			LOGGER.exception("Blocking command failed")
			self._conn.send({"type": "response", "id": msg_id, "error": str(exc)})

	# ------------------------------------------------------------------
	# Command handlers

	def _handle_initialize(self, dllPath: str, tibasePath: str, initialVoice: int = 1, **_kw) -> Dict:
		config = HostConfig(
			dll_path=dllPath,
			tibase_path=tibasePath,
			initial_voice=initialVoice,
		)
		self._runtime = SoftVoiceRuntime(self._conn, config)
		return self._runtime.start()

	def _handle_speak(self, text: str, **_kw) -> Dict:
		self._runtime.speak(text)
		return {"status": "ok"}

	def _handle_stop(self, **_kw) -> Dict:
		if self._runtime:
			self._runtime.stop()
		self._wait_for_speak_thread()
		return {"status": "ok"}

	def _handle_dll_call(self, funcName: str, value: int, **_kw) -> Dict:
		self._runtime.dll_call(funcName, value)
		return {"status": "ok"}

	def _handle_get_format(self, **_kw) -> Dict:
		return self._runtime.get_format()

	def _handle_delete(self, **_kw) -> Dict:
		self._wait_for_speak_thread()
		if self._runtime:
			self._runtime.delete()
		self._should_exit = True
		return {"status": "ok"}

	def _wait_for_speak_thread(self) -> None:
		if self._speak_thread and self._speak_thread.is_alive():
			self._speak_thread.join(timeout=30)
		self._speak_thread = None


def main() -> None:
	parser = argparse.ArgumentParser(description="SoftVoice 32-bit helper")
	parser.add_argument("--address", required=True)
	parser.add_argument("--authkey", required=True)
	parser.add_argument("--log-dir", default=None)
	args = parser.parse_args()

	configure_logging(args.log_dir)
	LOGGER.info("Connecting to controller at %s", args.address)

	host, port_str = args.address.split(":")
	address = (host, int(port_str))
	authkey = bytes.fromhex(args.authkey)
	conn = _ipc.connect_to_listener(address, authkey)
	controller = HostController(conn)
	controller.serve_forever()


if __name__ == "__main__":
	main()
