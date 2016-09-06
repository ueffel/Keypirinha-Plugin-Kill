import keypirinha_wintypes as kpwt
import ctypes as ct

psapi = ct.windll.psapi
ntdll = ct.windll.ntdll
kernel = kpwt.kernel32

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010
PROCESS_TERMINATE = 0x0001

TH32CS_SNAPPROCESS = 0x00000002
MAX_PATH_LEN = 260

class PROCESSENTRY32(ct.Structure):
    _fields_ = [('dwSize', ct.wintypes.DWORD),
                ('cntUsage', ct.wintypes.DWORD),
                ('th32ProcessID', ct.wintypes.DWORD),
                ('th32DefaultHeapID', ct.POINTER(ct.c_ulong)),
                ('th32ModuleID', ct.wintypes.DWORD),
                ('cntThreads', ct.wintypes.DWORD),
                ('th32ParentProcessID', ct.wintypes.DWORD),
                ('pcPriClassBase', ct.c_long),
                ('dwFlags', ct.wintypes.DWORD),
                ('szExeFile', ct.c_char * 260)]

def getpids():
    """
        Uses a process snapshot to the a list of all process with pid and name
    """
    proc_snapshot = kernel.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    pe32 = PROCESSENTRY32()
    pe32.dwSize = ct.c_ulong(ct.sizeof(PROCESSENTRY32))
    if not proc_snapshot:
        print(kernel.GetLastError())
        return []
    procs = []
    if kernel.Process32First(proc_snapshot, ct.byref(pe32)):
        procs.append((pe32.th32ProcessID, ct.string_at(pe32.szExeFile).decode()))
    while kernel.Process32Next(proc_snapshot, ct.byref(pe32)):
        procs.append((pe32.th32ProcessID, ct.string_at(pe32.szExeFile).decode()))
    return procs


class AccessDenied(Exception):
    """
        This exception is thrown when the process handle couldn't be obtained
    """
    pass

class Process(object):
    """
        Simple class as wrapper around some windows api functions
    """
    def __init__(self, pid):
        self._pid = pid
        self._exe = None
        self._cmdline = None
        self._init()

    def _init(self):
        """
            Tries to get some information about the process
        """
        # attempt 1
        proc_handle = kernel.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                                         False,
                                         self._pid)
        if proc_handle:
            self._init_default(proc_handle)
            kernel.CloseHandle(proc_handle)
            return

        # attempt 2
        proc_handle = kernel.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION,
                                         False,
                                         self._pid)
        if proc_handle:
            self._init_restricted(proc_handle)
            kernel.CloseHandle(proc_handle)
            return

    def _init_default(self, proc_handle):
        self._exe = self._get_full_path(proc_handle)
        self._cmdline = [self._exe]

    def _init_restricted(self, proc_handle):
        full_path = self._get_full_path(proc_handle)
        self._exe = full_path

    def _get_cmd_line(self, proc_handle):
        """
            Obtains all command line parameters of the process
        """

        # Hey nice, you're reading my code. Do you have an idea how to get the
        # command line of a process without calling ntdll.dll directly?
        # I would settle for something with wmi but without an external call to
        # wmic. Sadly there is no win32api, win32com or wmi module to import.
        # I'm out of ideas
        pass

    def _get_full_path(self, proc_handle):
        """
            Returns the full path of the processes executable image
            Calls QueryFullProcessImageName from kernel32.dll
        """
        image_name = ct.create_unicode_buffer(MAX_PATH_LEN)
        image_name_size = ct.wintypes.PDWORD(ct.c_ulong(MAX_PATH_LEN))
        if kernel.QueryFullProcessImageNameW(proc_handle, 0, image_name, image_name_size):
            return image_name[:image_name_size.contents.value]
        else:
            return ""

    def exe(self):
        """
            Returns the full path of the process if it was initialized
        """
        if self._exe:
            return self._exe
        else:
            raise AccessDenied

    def cmdline(self):
        """
            Returns a list with all command line parameters. First in the list
            is the executable path
        """
        if self._cmdline:
            return self._cmdline
        else:
            raise AccessDenied

    def kill(self):
        """
            Tries to kill the process with TerminateProcess from kernel32.dll
        """
        proc_handle = kernel.OpenProcess(PROCESS_TERMINATE, False, self._pid)
        if proc_handle:
            success = kernel.TerminateProcess(proc_handle, 1)
            kernel.CloseHandle(proc_handle)
            if not success:
                raise AccessDenied
        else:
            raise AccessDenied
