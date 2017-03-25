import keypirinha as kp
import keypirinha_util as kpu
import subprocess
import ctypes as ct

try:
    import comtypes.client as com_cl
except ImportError:
    com_cl = None

KERNEL = ct.windll.kernel32

PROCESS_TERMINATE = 0x0001

class Kill(kp.Plugin):
    """
        Plugin that lists running processes with name and commandline (if
        available) and kills the selected process(es)
    """

    def __init__(self):
        """
            Default constructor and initializing internal attributes
        """
        super().__init__()
        self._processes = []
        self._actions = []
        self._icons = {}
        self._default_action = "kill_by_id"
        # self._debug = True

    def on_events(self, flags):
        """
            Reloads the package config when its changed
        """
        if flags & kp.Events.PACKCONFIG:
            self._read_config()

    def _read_config(self):
        """
            Reads the default action from the config
        """
        settings = self.load_settings()

        possible_actions = [
            "kill_by_name",
            "kill_by_id",
            "kill_by_name_admin",
            "kill_by_id_admin"
        ]

        self._default_action = settings.get_enum(
            "default_action",
            "main",
            self._default_action,
            possible_actions
        )

    def on_start(self):
        """
            Creates the actions for killing the processes and register them
        """
        self._read_config()
        kill_by_name = self.create_action(
            name="kill_by_name",
            label="Kill by Name",
            short_desc="Kills all processes by that name"
        )
        self._actions.append(kill_by_name)

        kill_by_id = self.create_action(
            name="kill_by_id",
            label="Kill by PID",
            short_desc="Kills single process by its process id"
        )
        self._actions.append(kill_by_id)

        kill_by_name_admin = self.create_action(
            name="kill_by_name_admin",
            label="Kill by Name (as Admin)",
            short_desc="Kills all processes by that name"
            + " with elevated rights (taskkill /F /IM <exe>)"
        )
        self._actions.append(kill_by_name_admin)

        kill_by_id_admin = self.create_action(
            name="kill_by_id_admin",
            label="Kill by PID (as Admin)",
            short_desc="Kills single process by its process id"
            + " with elevated rights (taskkill /F /PID <pid>)"
        )
        self._actions.append(kill_by_id_admin)

        self.set_actions(kp.ItemCategory.KEYWORD, self._actions)

    def on_catalog(self):
        """
            Adds the kill command to the catalog
        """
        catalog = []

        killcmd = self.create_item(
            category=kp.ItemCategory.KEYWORD,
            label="Kill:",
            short_desc="Kills running processes",
            target="kill",
            args_hint=kp.ItemArgsHint.REQUIRED,
            hit_hint=kp.ItemHitHint.KEEPALL
        )

        catalog.append(killcmd)

        self.set_catalog(catalog)

    def _get_icon(self, source):
        """
            Tries to load the first icon within the source which should be a
            path to an executable
        """
        if source in self._icons:
            return self._icons[source]
        else:
            try:
                icon = self.load_icon("@{},0".format(source))
                self._icons[source] = icon
            except ValueError:
                self.dbg("Icon loading failed :( {}".format(source))
                icon = None
            return icon

    def on_activated(self):
        """
            Creates the list of running processes, when the Keypirinha Box is
            triggered
        """
        wmi = None
        if com_cl:
            wmi = com_cl.CoGetObject("winmgmts:")

        if wmi:
            self._get_processes_from_com_object(wmi)
        else:
            self.warn("Windows Management Service is not running.")
            self._get_processes_from_ext_call()

        self.dbg("%d running processes found" % len(self._processes))
        self.dbg("%d icons loaded" % len(self._icons))
        # self.dbg(self._icons)
        # for prc in self._processes:
        #     self.dbg(prc.raw_args())

    def _get_processes_from_com_object(self, wmi):
        """
            Creates the list of running processes
            Windows Management COMObject (WMI) to get the running processes
        """
        result_wmi = wmi.ExecQuery("SELECT ProcessId, Caption, Name, ExecutablePath, CommandLine "
                                   + "FROM Win32_Process")
        for proc in result_wmi:
            short_desc = ""
            if proc.Properties_["CommandLine"].Value is not None:
                short_desc = "(pid: {:5}) {}".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["CommandLine"].Value
                )
            elif proc.Properties_["ExecutablePath"].Value:
                short_desc = "(pid: {:5}) {}".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["ExecutablePath"].Value
                )
            elif proc.Properties_["Name"].Value:
                short_desc = "(pid: {:5}) {} ({})".format(
                    proc.Properties_["ProcessId"].Value,
                    proc.Properties_["Name"].Value,
                    "Probably only killable as admin or not at all"
                )
            item = self.create_item(
                category=kp.ItemCategory.KEYWORD,
                label=proc.Properties_["Caption"].Value,
                short_desc=short_desc,
                target=proc.Properties_["Name"].Value + "|"
                + str(proc.Properties_["ProcessId"].Value),
                icon_handle=self._get_icon(proc.Properties_["ExecutablePath"].Value),
                args_hint=kp.ItemArgsHint.REQUIRED,
                hit_hint=kp.ItemHitHint.IGNORE
            )
            self._processes.append(item)

    def _get_processes_from_ext_call(self):
        """
            FALLBACK
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
        for enc in ["cp437", "cp850", "cp1252", "utf8"]:
            try:
                output = output.replace(b"\r\r", b"\r")
                outstr = output.decode(enc)
                break
            except UnicodeDecodeError:
                self.dbg(enc + " threw exception")

        info = {}
        for line in outstr.splitlines():
            # self.dbg(line)
            if line.strip() == "":
                # build catalog item with gathered information from parsing
                if info and "Caption" in info:
                    short_desc = ""

                    if "CommandLine" in info and info["CommandLine"] != "":
                        short_desc = "(pid: {:5}) {}".format(
                            info["ProcessId"],
                            info["CommandLine"]
                        )
                    elif "ExecutablePath" in info and info["ExecutablePath"] != "":
                        short_desc = "(pid: {:5}) {}".format(
                            info["ProcessId"],
                            info["ExecutablePath"]
                        )
                    elif "Name" in info:
                        short_desc = "(pid: {:5}) {}".format(
                            info["ProcessId"],
                            info["Name"]
                        )

                    item = self.create_item(
                        category=kp.ItemCategory.KEYWORD,
                        label=info["Caption"],
                        short_desc=short_desc,
                        target=info["Name"] + "|" + info["ProcessId"],
                        icon_handle=self._get_icon(info["ExecutablePath"]),
                        args_hint=kp.ItemArgsHint.REQUIRED,
                        hit_hint=kp.ItemHitHint.IGNORE
                    )
                    self._processes.append(item)
                info = {}
            else:
                # Save key=value in info dict
                line_splitted = line.split("=")
                label = line_splitted[0]
                value = "=".join(line_splitted[1:])
                # Skip system processes that cant be killed
                if label == "Caption" and (value == "System Idle Process" or value == "System"):
                    continue
                info[label] = value

    def on_deactivated(self):
        """
            Emptys the process list, when Keypirinha Box is closed
        """
        self._processes = []

        # for ico in self._icons.values():
        #     ico.free()
        # self._icons = {}

    def on_suggest(self, user_input, items_chain):
        """
            Sets the list of running processes as suggestions
        """
        if not items_chain:
            return

        self.set_suggestions(self._processes, kp.Match.FUZZY, kp.Sort.SCORE_DESC)

    def on_execute(self, item, action):
        """
            Executes the selected (or default) kill action on the selected item
        """
        # get default action if no action was explicitly selected
        if action is None:
            for act in self._actions:
                if act.name() == self._default_action:
                    action = act


        if "_admin" in action.name():
            self._kill_process_admin(item.target(), action.name())
        else:
            self._kill_process_normal(item.target(), action.name())

    def _kill_process_normal(self, target, action_name):
        """
            Kills the selected process(es) using the windows api
        """
        target_name, target_pid = target.split("|")
        if "kill_by_name" in action_name:
            # loop over all processes and kill all by the same name
            for process_item in self._processes:
                pname, pid = process_item.target().split("|")
                if pname == target_name:
                    self.dbg("Killing process with id: {} and name: {}".format(pid, pname))
                    proc_handle = KERNEL.OpenProcess(PROCESS_TERMINATE, False, pid)
                    if not proc_handle:
                        self.warn("OpenProcess failed, ErrorCode: " + str(KERNEL.GetLastError()))
                        continue
                    success = KERNEL.TerminateProcess(proc_handle, 1)
                    if not success:
                        self.warn("TerminateProcess failed, ErrorCode: "
                                  + str(KERNEL.GetLastError()))
                        continue

        elif "kill_by_id" in action_name:
            # kill process with that pid
            self.dbg("Killing process with id: {} and name: {}".format(target_pid, target_name))
            pid = int(target_pid)
            proc_handle = KERNEL.OpenProcess(PROCESS_TERMINATE, False, pid)
            if not proc_handle:
                self.warn("OpenProcess failed, ErrorCode: " + str(KERNEL.GetLastError()))
                return
            success = KERNEL.TerminateProcess(proc_handle, 1)
            if not success:
                self.warn("TerminateProcess failed, ErrorCode: " + str(KERNEL.GetLastError()))
                return

    def _kill_process_admin(self, target, action_name):
        """
            Kills the selected process(es) using a call to windows' taskkill.exe
            with elevated rights
        """
        target_name, target_pid = target.split("|")
        args = ["taskkill", "/F"]

        # add parameters according to action
        if "kill_by_name" in action_name:
            args.append("/IM")
            # process name
            args.append(target_name)
        elif "kill_by_id" in action_name:
            args.append("/PID")
            # process id
            args.append(target_pid)

        self.dbg("Calling: {}".format(args))

        # show no window when executing
        kpu.shell_execute(args[0], args[1:], verb="runas", show=subprocess.SW_HIDE)
