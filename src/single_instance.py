"""Windows named-mutex guard for the GUI process."""

from __future__ import annotations

import ctypes
import os


GUI_MUTEX_NAME=r"Local\TVC_JOB_BOT_V080_GUI_SINGLE_INSTANCE"
ERROR_ALREADY_EXISTS=183


class _WindowsMutexApi:
    def __init__(self):
        if os.name!="nt":
            raise RuntimeError("Windows named mutex is only available on Windows")
        self.kernel32=ctypes.WinDLL("kernel32",use_last_error=True)
        self.kernel32.CreateMutexW.argtypes=(ctypes.c_void_p,ctypes.c_bool,ctypes.c_wchar_p)
        self.kernel32.CreateMutexW.restype=ctypes.c_void_p
        self.kernel32.CloseHandle.argtypes=(ctypes.c_void_p,)
        self.kernel32.CloseHandle.restype=ctypes.c_bool

    def create(self,name):
        ctypes.set_last_error(0)
        handle=self.kernel32.CreateMutexW(None,False,name)
        error=ctypes.get_last_error()
        if not handle:
            raise ctypes.WinError(error)
        return handle,error==ERROR_ALREADY_EXISTS

    def close(self,handle):
        if not self.kernel32.CloseHandle(handle):
            raise ctypes.WinError(ctypes.get_last_error())


class WindowsSingleInstance:
    """Own one named kernel object; Windows releases it if the process crashes."""

    def __init__(self,name=GUI_MUTEX_NAME,api=None):
        self.name=name
        self._api=api
        self._handle=None

    @property
    def acquired(self):
        return self._handle is not None

    def acquire(self):
        if self.acquired:
            return True
        api=self._api or _WindowsMutexApi()
        handle,already_exists=api.create(self.name)
        if already_exists:
            api.close(handle)
            return False
        self._api=api
        self._handle=handle
        return True

    def release(self):
        handle=self._handle
        if handle is None:
            return
        self._handle=None
        self._api.close(handle)

    def __enter__(self):
        if not self.acquire():
            raise RuntimeError("GUI instance already exists")
        return self

    def __exit__(self,_exc_type,_exc_value,_traceback):
        self.release()

