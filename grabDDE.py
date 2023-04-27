#grabDDE.py

import ctypes
import time

def set_global_constants(constants_dict):
    for key, value in constants_dict.items():
        globals()[key] = value

from ctypes import POINTER, WINFUNCTYPE, c_char_p, c_void_p, c_int, c_ulong, byref, create_string_buffer
from ctypes.wintypes import BOOL, DWORD, BYTE, INT, LPCWSTR, UINT, ULONG

HCONV= HDDEDATA = HSZ= c_void_p
LPBYTE, LPDWORD, LPSTR, ULONG_PTR = c_char_p, POINTER(DWORD), c_char_p, c_ulong
PCONVCONTEXT = c_void_p
formatx= {
    "DMLERR_NO_ERROR": 0,
    "CF_TEXT": 1,
    "CF_BITMAP": 2,
    "CF_METAFILEPICT": 3,
    "CF_SYLK": 4,
    "CF_DIF": 5,
    "CF_TIFF": 6,
    "CF_OEMTEXT": 7,
    "CF_DIB": 8,
    "CF_PALETTE": 9,
    "CF_PENDATA": 10,
    "CF_RIFF": 11,
    "CF_WAVE": 12,
    "CF_UNICODETEXT": 13,
    "CF_ENHMETAFILE": 14,
    "CF_HDROP": 15,
    "CF_LOCALE": 16,
    "CF_DIBV5": 17,
    "CF_MAX": 18
}
set_global_constants(formatx)

formaty = {
    "DDE_FACK": 0x8000,
    "DDE_FBUSY": 0x4000,
    "DDE_FDEFERUPD": 0x4000,
    "DDE_FACKREQ": 0x8000,
    "DDE_FRELEASE": 0x2000,
    "DDE_FREQUESTED": 0x1000,
    "DDE_FAPPSTATUS": 0x00FF,
    "DDE_FNOTPROCESSED": 0x0000,
    "DDE_FACKRESERVED": 0x8000 | 0x4000 | 0x00FF,
    "DDE_FADVRESERVED": 0x8000 | 0x4000,
    "DDE_FDATRESERVED": 0x8000 | 0x2000 | 0x1000,
    "DDE_FPOKRESERVED": 0x2000,
    "XTYPF_NOBLOCK": 0x0002,
    "XTYPF_NODATA": 0x0004,
    "XTYPF_ACKREQ": 0x0008,
    "XCLASS_MASK": 0xFC00,
    "XCLASS_BOOL": 0x1000,
    "XCLASS_DATA": 0x2000,
    "XCLASS_FLAGS": 0x4000,
    "XCLASS_NOTIFICATION": 0x8000,
    "XTYP_ERROR": 0x0000 | 0x8000 | 0x0002,
    "XTYP_ADVDATA": 0x0010 | 0x4000,
    "XTYP_ADVREQ": 0x0020 | 0x2000 | 0x0002,
    "XTYP_ADVSTART": 0x0030 | 0x1000,
    "XTYP_ADVSTOP": 0x0040 | 0x8000,
    "XTYP_EXECUTE": 0x0050 | 0x4000,
    "XTYP_CONNECT": 0x0060 | 0x1000 | 0x0002,
    "XTYP_CONNECT_CONFIRM": 0x0070 | 0x8000 | 0x0002,
    "XTYP_XACT_COMPLETE": 0x0080 | 0x8000,
    "XTYP_POKE": 0x0090 | 0x4000,
    "XTYP_REGISTER": 0x00A0 | 0x8000 | 0x0002,
    "XTYP_REQUEST": 0x00B0 | 0x2000,
    "XTYP_DISCONNECT": 0x00C0 | 0x8000 | 0x0002,
    "XTYP_UNREGISTER": 0x00D0 | 0x8000 | 0x0002,
    "XTYP_WILDCONNECT": 0x00E0 | 0x2000 | 0x0002,
    "XTYP_MONITOR": 0x00F0 | 0x8000 | 0x0002,
    "XTYP_MASK": 0x00F0,
    "XTYP_SHIFT": 4,
    "TIMEOUT_ASYNC": 0xFFFFFFFF,
}
set_global_constants(formaty)
def get_winfunc(libname, funcname, restype=None, argtypes=()):
    lib = windll.LoadLibrary(libname)
    func = getattr(lib, funcname)
    func.argtypes = argtypes
    func.restype = restype
    return func
DDECALLBACK = WINFUNCTYPE(HDDEDATA, UINT, UINT, HCONV, HSZ, HSZ, HDDEDATA, ULONG_PTR, ULONG_PTR)
class DDE(object):
    _funcs = {
        'DdeInitializeW': (None, 'DdeInitializeW', UINT, (POINTER(DWORD), DDECALLBACK, DWORD, DWORD)),
        'DdeCreateStringHandleW': (None, 'DdeCreateStringHandleW', HSZ, (DWORD, LPCWSTR, UINT)),
        'DdeUninitalize' : (None, 'DdeUninitialize', UINT, (DWORD,)),
        'DdeFreeStringHandle': (None, 'DdeFreeStringHandle', BOOL, (DWORD, HSZ)),
        'DdeConnect': (None, 'DdeConnect', HCONV, (DWORD, HSZ, HSZ, PCONVCONTEXT)),
        'DdeAccessData': (None, 'DdeAccessData', LPBYTE, (HDDEDATA, LPDWORD)),
        'DdeQueryStringW': (None, 'DdeQueryStringW', DWORD, (DWORD, HSZ, LPCWSTR, DWORD, UINT)),
        'DdeDisconnect': (None, 'DdeDisconnect', BOOL, (HCONV,)),
        'DdeLastError': (None, 'DdeLastError', UINT, (DWORD,)),
        'DdeClientTransaction': (None, 'DdeClientTransaction', DWORD, (LPBYTE, DWORD, HCONV, HSZ, UINT, UINT, DWORD, LPDWORD)),
    }
    def __init__(self):
        self.user32 = ctypes.windll.user32
        self.dde = ctypes.windll.LoadLibrary("user32")
        for name, (lib, funcname, restype, argtypes) in self._funcs.items():
            try:
                func = getattr(self.dde, funcname)
            except AttributeError:
                continue
            func.restype = restype
            func.argtypes = argtypes
            setattr(self, name, func)
    def __del__(self):
        del self.user32
        del self.dde

class DDEError(RuntimeError):
    def __init__(self, msg, idInst=None):
        RuntimeError.__init__(self, msg if idInst is None else f"{msg} (err={hex(self.dde.DdeLastError(idInst))})")

class DDEClient(object):
    def __init__(self, service, topic):
        self.dde = DDE()
        self._idInst = DWORD(0)
        self._hConv = HCONV(None)
        self._callback = DDECALLBACK(self._callback)
        if self.dde.DdeInitializeW(byref(self._idInst), self._callback, 0x00000010, 0) != DMLERR_NO_ERROR:
        # (res := self.dde.DdeInitialize(byref(self._idInst), self._callback, 0x00000010, 0)) != DMLERR_NO_ERROR:
            raise DDEError(f"Unable to register with DDEML (err={hex(res)})")
        hszService, hszTopic = self.dde.DdeCreateStringHandleW(self._idInst, service, 1200), self.dde.DdeCreateStringHandleW(self._idInst, topic, 1200)
        self._hConv = self.dde.DdeConnect(self._idInst, hszService, hszTopic, PCONVCONTEXT())
        for hsz in (hszTopic, hszService):
            self.dde.DdeFreeStringHandle(self._idInst, hsz)
        if not self._hConv:
            raise DDEError("Unable to establish a conversation with server", self._idInst)

    def __del__(self):
        try:
            if self._hConv:
                self.dde.DdeDisconnect(self._hConv)
            if self._idInst:
                self.dde.DdeUninitialize(self._idInst)
        except AttributeError:
            pass

        # self.dde.DdeUninitialize(self.dwDDEInstance)
        # if self._hConv: self.dde.DdeDisconnect(self._hConv)
        # if self._idInst: self.dde.DdeUninitialize(self._idInst)
        # self._hConv = None

    def advise(self, item, stop=False):
        hszItem = self.dde.DdeCreateStringHandleW(self._idInst, item, 1200)
        hDdeData = self.dde.DdeClientTransaction(LPBYTE(), 0, self._hConv, hszItem, CF_TEXT, XTYP_ADVSTOP if stop else XTYP_ADVSTART, TIMEOUT_ASYNC, LPDWORD())
        self.dde.DdeFreeStringHandle(self._idInst, hszItem)
        if not hDdeData:
            raise DDEError(f"Unable to {'stop' if stop else 'start'} advise", self._idInst)
        self.dde.DdeFreeStringHandle(self._idInst, hDdeData)

    def execute(self, command, timeout=5000):
        hDdeData = self.dde.DdeClientTransaction(c_char_p(command), DWORD(len(command) + 1), self._hConv, HSZ(), CF_TEXT, XTYP_EXECUTE, timeout, LPDWORD())
        if not hDdeData:
            raise DDEError("Unable to send command", self._idInst)
        self.dde.DdeFreeDataHandle(hDdeData)

    def query_string(self, hsz, max_length=128, code_page=1200):
        buffer = create_unicode_buffer(max_length)
        result = self.dde.DdeQueryStringW(self._idInst, hsz, buffer, max_length, code_page)
        if result == 0:
            raise DDEError("Unable to query string", self._idInst)
        return buffer.value

    def request(self, item, timeout=5000):
        hszItem = self.dde.DdeCreateStringHandleW(self._idInst, item, 1200)
        hDdeData = self.dde.DdeClientTransaction(LPBYTE(), 0, self._hConv, hszItem, CF_TEXT, XTYP_REQUEST, timeout, LPDWORD())
        self.dde.DdeFreeStringHandle(self._idInst, hszItem)
        if not hDdeData:
            raise DDEError("Unable to request item", self._idInst)

        if timeout != TIMEOUT_ASYNC:
            try:
                pData = self.dde.DdeAccessData(hDdeData, DWORD(0))
            except:
                pData = None

        if not pData:
            time.sleep(0.05)
            pData = self.request(item)
        return pData

    def callback(self, value, item=None):
        print(f"{item}: {value}")

    def _callback(self, wType, uFmt, hConv, hsz1, hsz2, hDdeData, dwData1, dwData2):
        dwSize = DWORD(0)
        pData = self.dde.DdeAccessData(hDdeData, byref(dwSize))
        if pData:
            item = create_string_buffer(b'\000' * 128)
            self.dde.DdeQueryStringW(self._idInst, hsz2, item, 128, 1004)
            self.callback(pData, item.value)
            self.dde.DdeUnaccessData(hDdeData)
        return DDE_FACK

    def WinMSGLoop():
        LPMSG = POINTER(MSG)
        LRESULT = c_ulong
        GetMessage = get_winfunc("user32", "GetMessageW", BOOL, (LPMSG, HWND, UINT, UINT))
        TranslateMessage = get_winfunc("user32", "TranslateMessage", BOOL, (LPMSG,))
        DispatchMessage = get_winfunc("user32", "DispatchMessageW", LRESULT, (LPMSG,))

        msg = MSG()
        lpmsg = byref(msg)
        while GetMessage(lpmsg, HWND(), 0, 0) > 0:
            TranslateMessage(lpmsg)
            DispatchMessage(lpmsg)
