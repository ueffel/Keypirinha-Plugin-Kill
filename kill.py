from .lib.alttab import AltTab
import keypirinha as kp
import keypirinha_util as kpu
import subprocess
import ctypes as ct
import time
import traceback
import asyncio

try:
    import comtypes.client as com_cl
except ImportError:
    com_cl = None

KERNEL = ct.windll.kernel32

CommandLineToArgvW = ct.windll.shell32.CommandLineToArgvW
CommandLineToArgvW.argtypes = [ct.wintypes.LPCWSTR, ct.POINTER(ct.c_int)]
CommandLineToArgvW.restype = ct.POINTER(ct.wintypes.LPWSTR)

PostMessageW = ct.windll.user32.PostMessageW
PostMessageW.argtypes = [ct.wintypes.HWND, ct.c_uint, ct.wintypes.WPARAM, ct.wintypes.LPARAM]
PostMessageW.restype = ct.c_long

PROCESS_TERMINATE = 0x0001
SYNCHRONIZE = 0x00100000
WAIT_OBJECT_0 = 0x00000000
WAIT_TIMEOUT = 0x00000102
WAIT_FAILED = 0xFFFFFFFF
WM_CLOSE = 0x0010
RESTARTABLE = kp.ItemCategory.USER_BASE + 1


class Kill(kp.Plugin):
    """Plugin that lists running processes with name and commandline (if available) and kills the selected process(es)
    """
    ACTION_KILL_BY_ID = "kill_by_id"
    ACTION_KILL_BY_NAME = "kill_by_name"
    ACTION_KILL_RESTART_BY_ID = "kill_and_restart_by_id"
    ADMIN_SUFFIX = "_admin"
    ACTION_KILL_BY_ID_ADMIN = ACTION_KILL_BY_ID + ADMIN_SUFFIX
    ACTION_KILL_BY_NAME_ADMIN = ACTION_KILL_BY_NAME + ADMIN_SUFFIX
    DEFAULT_ITEM_LABEL = "Kill:"

    def __init__(self):
        """Default constructor and initializing internal attributes
        """
        super().__init__()
        self._processes = []
        self._processes_with_window = {}
        self._actions = []
        self._icons = {}
        self._default_action = self.ACTION_KILL_BY_ID
        self._hide_background = False
        self._default_icon = None
        self._item_label = self.DEFAULT_ITEM_LABEL
        self.__executing = False

    def on_events(self, flags):
        """Reloads the package config when its changed
        """
        if flags & kp.Events.PACKCONFIG:
            self._read_config()

    def _read_config(self):
        """Reads the config
        """
        self.dbg("Reading config")
        settings = self.load_settings()

        self._debug = settings.get_bool("debug", "main", False)

        possible_actions = [
            self.ACTION_KILL_BY_NAME,
            self.ACTION_KILL_BY_ID,
            self.ACTION_KILL_BY_NAME_ADMIN,
            self.ACTION_KILL_BY_ID_ADMIN
        ]

        self._default_action = settings.get_enum(
            "default_action",
            "main",
            self._default_action,
            possible_actions
        )
        self.dbg("default_action =", self._default_action)

        self._hide_background = settings.get_bool("hide_background", "main", False)
        self.dbg("hide_background =", self._hide_background)

        self._item_label = settings.get("item_label", "main", self.DEFAULT_ITEM_LABEL)
        self.dbg("item_label =", self._item_label)

    def on_start(self):
        """Reads the config, creates the actions for killing the processes and register them
        """
        self._read_config()

        kill_by_name = self.create_action(
            name=self.ACTION_KILL_BY_NAME,
            label="Kill by Name",
            short_desc="Kills all processes by that name"
        )
        self._actions.append(kill_by_name)

        kill_by_id = self.create_action(
            name=self.ACTION_KILL_BY_ID,
            label="Kill by PID",
            short_desc="Kills single process by its process id"
        )
        self._actions.append(kill_by_id)

        kill_by_name_admin = self.create_action(
            name=self.ACTION_KILL_BY_NAME_ADMIN,
            label="Kill by Name (as Admin)",
            short_desc="Kills all processes by that name"
            + " with elevated rights (taskkill /F /IM <exe>)"
        )
        self._actions.append(kill_by_name_admin)

        kill_by_id_admin = self.create_action(
            name=self.ACTION_KILL_BY_ID_ADMIN,
            label="Kill by PID (as Admin)",
            short_desc="Kills single process by its process id"
            + " with elevated rights (taskkill /F /PID <pid>)"
        )
        self._actions.append(kill_by_id_admin)

        self.set_actions(kp.ItemCategory.KEYWORD, self._actions)

        kill_and_restart_by_id = self.create_action(
            name=self.ACTION_KILL_RESTART_BY_ID,
            label="Kill by PID and restart application",
            short_desc="Kills single process by its process id"
            + " and tries to restart it"
        )

        self._actions.append(kill_and_restart_by_id)
        self.set_actions(RESTARTABLE, self._actions)

        self._default_icon = self.load_icon("res://{}/kill.ico".format(self.package_full_name()))

    def on_catalog(self):
        """Adds the kill command to the catalog
        """
        catalog = []
        killcmd = self.create_item(
            category=kp.ItemCategory.KEYWORD,
            label=self._item_label,
            short_desc="Kills running processes",
            target="kill",
            args_hint=kp.ItemArgsHint.REQUIRED,
            hit_hint=kp.ItemHitHint.KEEPALL
        )
        catalog.append(killcmd)
        self.set_catalog(catalog)

    def _get_icon(self, source):
        """Tries to load the first icon within the source which should be a path to an executable
        """
        if not source:
            return self._default_icon

        if source in self._icons:
            return self._icons[source]
        else:
            try:
                icon = self.load_icon("@{},0".format(source))
                self._icons[source] = icon
            except ValueError:
                self.dbg("Icon loading failed :(", source)
                icon = None
            if not icon:
                return self._default_icon
            return icon

    def _get_processes(self):
        """Creates the list of running processes, when the Keypirinha Box is triggered
        """
        start_time = time.time()

        wmi = None
        if com_cl:
            wmi = com_cl.CoGetObject("winmgmts:")

        if wmi:
            self._get_processes_from_com_object(wmi)
        else:
            self.warn("Windows Management Service is not running.")
            self._get_processes_from_ext_call()

        elapsed = time.time() - start_time

        self.info("Found {} running processes in {:0.1f} seconds".format(len(self._processes), elapsed))
        self.dbg(len(self._icons), "icons loaded")

    def _get_windows(self):
        """Gets the list of open windows create a mapping between pid and hwnd
        """
        self.dbg("Getting windows")
        try:
            handles = AltTab.list_alttab_windows()
        except OSError:
            self.err("Failed to list windows.", traceback.format_exc())
            return

        self._processes_with_window = {}

        for hwnd in handles:
            try:
                _, proc_id = AltTab.get_window_thread_process_id(hwnd)
                if proc_id in self._processes_with_window:
                    self._processes_with_window[proc_id].append(hwnd)
                else:
                    self._processes_with_window[proc_id] = [hwnd]
            except OSError:
                continue
        self.dbg(len(self._processes_with_window), "windows found")

    def _get_processes_from_com_object(self, wmi):
        """Creates the list of running processes

        Uses Windows Management COMObject (WMI) to get the running processes
        """
        result_wmi = wmi.ExecQuery("SELECT ProcessId, Caption, Name, ExecutablePath, CommandLine "
                                   "FROM Win32_Process")
        for proc in result_wmi:
            pid = proc.Properties_["ProcessId"].Value
            is_foreground = pid in self._processes_with_window
            if is_foreground:
                window_title = AltTab.get_window_text(self._processes_with_window[pid][0])
            else:
                window_title = ""

            if self._hide_background and not is_foreground:
                continue

            short_desc = ""
            category = kp.ItemCategory.KEYWORD
            databag = {}
            if proc.Properties_["CommandLine"].Value:
                short_desc = "(pid: {:>5}) {}".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["CommandLine"].Value
                )
                category = RESTARTABLE
                databag["CommandLine"] = proc.Properties_["CommandLine"].Value
            elif proc.Properties_["ExecutablePath"].Value:
                short_desc = "(pid: {:>5}) {}".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["ExecutablePath"].Value
                )
            elif proc.Properties_["Name"].Value:
                short_desc = "(pid: {:>5}) {} ({})".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["Name"].Value,
                    "Probably only killable as admin or not at all"
                )

            if proc.Properties_["ExecutablePath"].Value:
                databag["ExecutablePath"] = proc.Properties_["ExecutablePath"].Value

            if not self._hide_background:
                if is_foreground:
                    label = '{}: "{}" ({})'.format(
                        proc.Properties_["Caption"].Value,
                        window_title,
                        'foreground'
                    )
                else:
                    label = '{} ({})'.format(proc.Properties_["Caption"].Value, 'background')
            else:
                label = '{}: "{}"'.format(
                    proc.Properties_["Caption"].Value,
                    window_title
                )

            item = self.create_item(
                category=category,
                label=label,
                short_desc=short_desc,
                target=proc.Properties_["Name"].Value + "|" + str(proc.Properties_["ProcessId"].Value),
                icon_handle=self._get_icon(proc.Properties_["ExecutablePath"].Value),
                args_hint=kp.ItemArgsHint.FORBIDDEN,
                hit_hint=kp.ItemHitHint.IGNORE,
                data_bag=str(databag)
            )
            self._processes.append(item)

    def _get_processes_from_ext_call(self):
        """FALLBACK

        Creates the list of running processes
        Uses Windows' "wmic.exe" tool to get the running processes
        """
        # Using external call to wmic to get the list of running processes
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output, err = subprocess.Popen(["wmic",
                                        "process",
                                        "get",
                                        "ProcessId,Caption,",
                                        "Name,ExecutablePath,CommandLine",
                                        "/FORMAT:LIST"],
                                       stdout=subprocess.PIPE,
                                       # universal_newlines=True,
                                       shell=False,
                                       startupinfo=startupinfo).communicate()
        # log error if any
        if err:
            self.err(err)

        # Parsing process list from output
        outstr = None
        for enc in ["cp437", "cp850", "cp1252", "utf8"]:
            try:
                output = output.replace(b"\r\r", b"\r")
                outstr = output.decode(enc)
                break
            except UnicodeDecodeError:
                self.dbg(enc, "threw exception")

        if not outstr:
            self.warn("decoding of output failed")
            return

        info = {}
        for line in outstr.splitlines():
            if line.strip() == "":
                # build catalog item with gathered information from parsing
                if info and "Caption" in info:
                    is_foreground = int(info["ProcessId"]) in self._processes_with_window
                    if self._hide_background and not is_foreground:
                        continue

                    short_desc = ""
                    category = kp.ItemCategory.KEYWORD
                    databag = {}
                    if "CommandLine" in info and info["CommandLine"] != "":
                        short_desc = "(pid: {:>5}) {}".format(
                            info["ProcessId"],
                            info["CommandLine"]
                        )
                        category = RESTARTABLE
                        databag["CommandLine"] = info["CommandLine"]
                    elif "ExecutablePath" in info and info["ExecutablePath"] != "":
                        short_desc = "(pid: {:>5}) {}".format(
                            info["ProcessId"],
                            info["ExecutablePath"]
                        )
                    elif "Name" in info:
                        short_desc = "(pid: {:>5}) {}".format(
                            info["ProcessId"],
                            info["Name"]
                        )

                    if "ExecutablePath" in info and info["ExecutablePath"] != "":
                        databag["ExecutablePath"] = info["ExecutablePath"]

                    label = info["Caption"]
                    if not self._hide_background:
                        if is_foreground:
                            label = "{} (foreground)".format(label)
                        else:
                            label = "{} (background)".format(label)

                    item = self.create_item(
                        category=category,
                        label=label,
                        short_desc=short_desc,
                        target=info["Name"] + "|" + info["ProcessId"],
                        icon_handle=self._get_icon(info["ExecutablePath"]),
                        args_hint=kp.ItemArgsHint.FORBIDDEN,
                        hit_hint=kp.ItemHitHint.IGNORE,
                        data_bag=str(databag)
                    )
                    self._processes.append(item)
                info = {}
            else:
                # Save key=value in info dict
                line_splitted = line.split("=")
                label = line_splitted[0]
                value = "=".join(line_splitted[1:])
                # Skip system processes that cant be killed
                if label == "Caption" and value in ("System Idle Process", "System"):
                    continue
                info[label] = value

    def _is_running(self, pid):
        wmi = None
        if com_cl:
            wmi = com_cl.CoGetObject("winmgmts:")

        if wmi:
            return self._is_running_from_com_object(wmi, pid)
        else:
            return self._is_running_from_ext_call(pid)

    def _is_running_from_com_object(self, wmi, pid):
        result_wmi = wmi.ExecQuery("SELECT ProcessId, Caption, Name, ExecutablePath, CommandLine "
                                   "FROM Win32_Process "
                                   "WHERE ProcessId = {}".format(pid))
        running = len(result_wmi) > 0
        self.dbg("(wmi) process with id ", pid, "running" if running else "not running")
        return running

    def _is_running_from_ext_call(self, pid):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output, err = subprocess.Popen(["wmic",
                                        "process",
                                        "where",
                                        "ProcessId={}".format(pid),
                                        "get",
                                        "ProcessId",
                                        "/FORMAT:LIST"],
                                       stdout=subprocess.PIPE,
                                       # universal_newlines=True,
                                       shell=False,
                                       startupinfo=startupinfo).communicate()
        # log error if any
        if err:
            self.err(err)

        # Parsing process list from output
        outstr = None
        for enc in ["cp437", "cp850", "cp1252", "utf8"]:
            try:
                output = output.replace(b"\r\r", b"\r")
                outstr = output.decode(enc)
                break
            except UnicodeDecodeError:
                self.dbg(enc, "threw exception")

        if not outstr:
            self.warn("decoding of output failed")
            return False

        running = "ProcessId={}".format(pid) in outstr.splitlines()
        self.dbg("(wmic) process with id ", pid, "running" if running else "not running")
        return running

    def on_deactivated(self):
        """Cleans up, when Keypirinha Box is closed
        """
        if not self.__executing:
            self._cleanup()

    def _cleanup(self):
        """Empties the process list, window list and frees the icon handles
        """
        self.dbg("Cleaning up")
        self._processes_with_window = {}
        self._processes = []

        for ico in self._icons.values():
            ico.free()
        self._icons = {}

    def on_suggest(self, user_input, items_chain):
        """Sets the list of running processes as suggestions
        """
        if not items_chain:
            return

        if not self._processes_with_window:
            self._get_windows()

        if not self._processes:
            self._get_processes()

        self.set_suggestions(self._processes, kp.Match.FUZZY, kp.Sort.SCORE_DESC)

    def on_execute(self, item, action):
        """Executes the selected (or default) kill action on the selected item
        """
        self.__executing = True
        loop = None
        try:
            # get default action if no action was explicitly selected
            if action is None:
                for act in self._actions:
                    if act.name() == self._default_action:
                        action = act

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            if action.name().endswith(self.ADMIN_SUFFIX):
                self._kill_process_admin(item, action.name())
            else:
                killing_task = asyncio.ensure_future(self._kill_process_normal(item, action.name()))
                loop.run_until_complete(killing_task)
        finally:
            self._cleanup()
            self.__executing = False
            if loop:
                loop.close()

    async def _kill_process_normal(self, target_item, action_name):
        """Kills the selected process(es) using the windows api
        """
        target_name, target_pid = target_item.target().split("|")
        if action_name.startswith(self.ACTION_KILL_BY_NAME):
            # loop over all processes and kill all by the same name
            kill_tasks = {}
            for process_item in self._processes:
                pname, pid = process_item.target().split("|")
                pid = int(pid)
                if pname == target_name:
                    self.dbg("Killing process with id: {} and name: {}".format(pid, pname))
                    kill_tasks[pid] = asyncio.get_event_loop().run_in_executor(None, self._kill_by_pid, pid)
            await asyncio.gather(*kill_tasks.values(), return_exceptions=True)

            self.dbg("Kill tasks finished")
            for pid, kill_task in kill_tasks.items():
                exc = kill_task.exception()
                if exc:
                    self.err(exc)
                    self.dbg(traceback.format_exception(exc.__class__, exc, exc.__traceback__))
                    continue
                result = kill_task.result()
                if not result:
                    self.warn("Killing process with pid", pid, "failed")

        elif action_name.startswith(self.ACTION_KILL_BY_ID):
            # kill process with that pid
            self.dbg("Killing process with id: {} and name: {}".format(target_pid, target_name))
            pid = int(target_pid)
            killed = await asyncio.get_event_loop().run_in_executor(None, self._kill_by_pid, pid)
            if not killed:
                self.warn("Killing process with id", pid, "failed")
        elif self.ACTION_KILL_RESTART_BY_ID:
            # kill process with that pid and try to restart it
            self.dbg("Killing process with id: {} and name: {}".format(target_pid, target_name))
            pid = int(target_pid)
            killed = await asyncio.get_event_loop().run_in_executor(None,
                                                                    lambda: self._kill_by_pid(pid, wait_for_exit=True))
            if not killed:
                self.warn("Killing process with id", pid, "failed. Not restarting")
                return
            databag = eval(target_item.data_bag())
            self.dbg("databag for process: ", databag)
            if "CommandLine" not in databag:
                self.warn("No commandline, cannot restart")
                return

            cmd = ct.wintypes.LPCWSTR(databag["CommandLine"])
            argc = ct.c_int(0)
            argv = CommandLineToArgvW(cmd, ct.byref(argc))
            if argc.value <= 0:
                self.dbg("No args parsed")
                return

            args = [argv[i] for i in range(0, argc.value)]
            self.dbg("CommandLine args from CommandLineToArgvW:", args)
            if args[0] == "" or args[0].isspace():
                args[0] = databag["ExecutablePath"]
            self.dbg("Restarting:", args)
            kpu.shell_execute(args[0], args[1:])

    def _kill_by_pid(self, pid, wait_for_exit=False):
        proc_handle = KERNEL.OpenProcess(PROCESS_TERMINATE | SYNCHRONIZE, False, pid)
        if not proc_handle:
            self.dbg("OpenProcess failed, ErrorCode:", KERNEL.GetLastError())
            return False

        if pid in self._processes_with_window:
            self.dbg("Posting WM_CLOSE to", len(self._processes_with_window[pid]), "windows")
            for hwnd in self._processes_with_window[pid]:
                success = PostMessageW(hwnd, ct.c_uint(WM_CLOSE), 0, 0)
                self.dbg("PostMessageW return:", success)

            self.dbg("Waiting for exit")
            timeout = ct.wintypes.DWORD(5000)
            result = KERNEL.WaitForSingleObject(proc_handle, timeout)
            if result == WAIT_OBJECT_0:
                self.dbg("process exited clean.")
                return True
            if result == WAIT_TIMEOUT:
                self.dbg("WaitForSingleObject timed out.")
            else:
                self.warn("Something weird happened in WaitForSingleObject:", result)
            self.dbg("ErrorCode:", KERNEL.GetLastError())
            if not self._is_running(pid):
                return True

        self.dbg("TerminateProcess!")
        success = KERNEL.TerminateProcess(proc_handle, 1)
        if not success:
            self.warn("TerminateProcess failed, ErrorCode:", KERNEL.GetLastError())
            return False

        if wait_for_exit:
            self.dbg("Waiting for exit")
            timeout = ct.wintypes.DWORD(1000)
            result = KERNEL.WaitForSingleObject(proc_handle, timeout)
            if result == WAIT_FAILED:
                self.warn("WaitForSingleObject failed, ErrorCode:", KERNEL.GetLastError())
                return False
            if result == WAIT_TIMEOUT:
                self.warn("WaitForSingleObject timed out.")
                return False
            if result != WAIT_OBJECT_0:
                self.warn("Something weird happened in WaitForSingleObject:", result)
                return False

        return True

    def _kill_process_admin(self, target_item, action_name):
        """Kills the selected process(es) using a call to windows' taskkill.exe  with elevated rights
        """
        target_name, target_pid = target_item.target().split("|")
        args = ["taskkill", "/F"]

        # add parameters according to action
        if action_name.startswith(self.ACTION_KILL_BY_NAME):
            args.append("/IM")
            # process name
            args.append(target_name)
        elif action_name.startswith(self.ACTION_KILL_BY_ID):
            args.append("/PID")
            # process id
            args.append(target_pid)

        self.dbg("Calling:", args)
        kpu.shell_execute(args[0], args[1:], verb="runas", show=subprocess.SW_HIDE)
