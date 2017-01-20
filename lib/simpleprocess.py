import ctypes as ct
from ctypes import wintypes as wt

NTDLL = ct.windll.ntdll
KERNEL = ct.windll.kernel32

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010
PROCESS_TERMINATE = 0x0001

TH32CS_SNAPPROCESS = 0x00000002
MAX_PATH_LEN = 260

class PROCESSENTRY32(ct.Structure):
    _fields_ = (('dwSize', wt.DWORD),
                ('cntUsage', wt.DWORD),
                ('th32ProcessID', wt.DWORD),
                ('th32DefaultHeapID', ct.POINTER(ct.c_ulong)),
                ('th32ModuleID', wt.DWORD),
                ('cntThreads', wt.DWORD),
                ('th32ParentProcessID', wt.DWORD),
                ('pcPriClassBase', ct.c_long),
                ('dwFlags', wt.DWORD),
                ('szExeFile', ct.c_char * 260))

class UNICODE_STRING(ct.Structure):
    _fields_ = (("Length", wt.USHORT),
                ("MaximumLength", wt.USHORT),
                ("Buffer", ct.POINTER(ct.c_wchar_p)))

class RTL_USER_PROCESS_PARAMETERS(ct.Structure):
    _fields_ = (("Reserved1", wt.BYTE*16),
                ("Reserved2", wt.LPVOID*10),
                ("ImagePathName", UNICODE_STRING),
                ("CommandLine", UNICODE_STRING))

class PEB(ct.Structure):
    _fields_ = (("Reserved1", wt.BYTE*2),
                ("BeingDebugged", wt.BYTE),
                ("Reserved2", wt.BYTE),
                ("Reserved3", wt.LPVOID*2),
                ("Ldr", wt.LPVOID),
                ("ProcessParameters", ct.POINTER(RTL_USER_PROCESS_PARAMETERS)),
                ("Reserved4", wt.BYTE*104),
                ("Reserved5", wt.LPVOID*52),
                ("PostProcessInitRoutine", wt.LPVOID),
                ("Reserved6", wt.BYTE*128),
                ("Reserved7", wt.LPVOID),
                ("SessionId", wt.ULONG))

class PROCESS_BASIC_INFORMATION(ct.Structure):
    _fields_ = (("Reserved1", wt.LPVOID),
                ("PebBaseAddress", ct.POINTER(PEB)),
                ("Reserved2", wt.LPVOID*2),
                ("UniqueProcessId", ct.POINTER(wt.ULONG)),
                ("Reserved3", wt.LPVOID))

def getpids():
    """
        Uses a process snapshot to the a list of all process with pid and name
    """
    proc_snapshot = KERNEL.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    pe32 = PROCESSENTRY32()
    pe32.dwSize = ct.c_ulong(ct.sizeof(PROCESSENTRY32))
    if not proc_snapshot:
        print(KERNEL.GetLastError())
        return []
    procs = []
    if KERNEL.Process32First(proc_snapshot, ct.byref(pe32)):
        procs.append((pe32.th32ProcessID, ct.string_at(pe32.szExeFile).decode()))
    while KERNEL.Process32Next(proc_snapshot, ct.byref(pe32)):
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
        proc_handle = KERNEL.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                                         False,
                                         self._pid)
        if proc_handle:
            self._init_default(proc_handle)
            KERNEL.CloseHandle(proc_handle)
            return

        # attempt 2
        proc_handle = KERNEL.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION,
                                         False,
                                         self._pid)
        if proc_handle:
            self._init_restricted(proc_handle)
            KERNEL.CloseHandle(proc_handle)
            return

    def _init_default(self, proc_handle):
        self._exe = self._get_full_path(proc_handle)
        self._cmdline = self._get_cmd_line(proc_handle)

    def _init_restricted(self, proc_handle):
        full_path = self._get_full_path(proc_handle)
        self._exe = full_path

    def _get_cmd_line(self, proc_handle):
        """
            Obtains command line parameters of the process if possible
        """
        pbi = PROCESS_BASIC_INFORMATION()
        peb = PEB()
        rtl = RTL_USER_PROCESS_PARAMETERS()

        # getting the basic process infos including the process enviroment block (PEB) adress
        NTDLL.NtQueryInformationProcess(proc_handle, 0, ct.byref(pbi), ct.sizeof(pbi), None)

        # getting the PEB
        bytes_read = ct.c_size_t()
        success = KERNEL.ReadProcessMemory(proc_handle,
                                           pbi.PebBaseAddress,
                                           ct.byref(peb),
                                           ct.sizeof(PEB),
                                           ct.byref(bytes_read))
        if not success:
            return None

        # getting the process parameter infos
        success = KERNEL.ReadProcessMemory(proc_handle,
                                           peb.ProcessParameters,
                                           ct.byref(rtl),
                                           ct.sizeof(RTL_USER_PROCESS_PARAMETERS),
                                           ct.byref(bytes_read))
        if not success:
            return None

        # getting the command line
        cmdline = (ct.c_wchar * rtl.CommandLine.Length)()
        if rtl.CommandLine.Buffer:
            KERNEL.ReadProcessMemory(proc_handle,
                                     rtl.CommandLine.Buffer,
                                     ct.byref(cmdline),
                                     ct.sizeof(cmdline),
                                     ct.byref(bytes_read))
            return cmdline.value
        else:
            return None


    def _get_full_path(self, proc_handle):
        """
            Returns the full path of the processes executable image
            Calls QueryFullProcessImageName from kernel32.dll
        """
        image_name = ct.create_unicode_buffer(MAX_PATH_LEN)
        image_name_size = wt.PDWORD(ct.c_ulong(MAX_PATH_LEN))
        if KERNEL.QueryFullProcessImageNameW(proc_handle, 0, image_name, image_name_size):
            return image_name[:image_name_size.contents.value]
        else:
            return None

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
        proc_handle = KERNEL.OpenProcess(PROCESS_TERMINATE, False, self._pid)
        if proc_handle:
            success = KERNEL.TerminateProcess(proc_handle, 1)
            KERNEL.CloseHandle(proc_handle)
            if not success:
                raise AccessDenied
        else:
            raise AccessDenied

# for (pid, name) in getpids():
#     proc = Process(pid)
#     try:
#         print("{:5d} {:20} {}".format(pid, name, proc.cmdline()))
#     except AccessDenied:
#         print("{:5d} {:20}".format(pid, name))
