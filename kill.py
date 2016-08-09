import keypirinha as kp
import keypirinha_util as kpu
import subprocess

class Kill(kp.Plugin):
    """
        Plugin that lists running processes with name and commandline (if
        available) and kills the select process
    """

    def __init__(self):
        """
            Default constructor and initializing internal attributes
        """
        super().__init__()
        self._processes = []
        self._actions = []
        self._icons = {}
        self._default_action = "kill_by_name"
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
            "kill_by_name",
            possible_actions
        )

    def on_start(self):
        """
            Creates the actions for killing the processes and register them
        """
        self._read_config()
        kill_by_name = self.create_action(
            name="kill_by_name",
            label="Kill all processes by that name",
            short_desc="Kill by Name (taskkill /F /IM <exe>)"
        )
        self._actions.append(kill_by_name)

        kill_by_id = self.create_action(
            name="kill_by_id",
            label="Kill single process",
            short_desc="Kill by PID (taskkill /F /PID <pid>)"
        )
        self._actions.append(kill_by_id)

        kill_by_name_admin = self.create_action(
            name="kill_by_name_admin",
            label="Kill all processes by that name (as Admin)",
            short_desc="Kill by Name with elevated rights (taskkill /F /IM <exe>)"
        )
        self._actions.append(kill_by_name_admin)

        kill_by_id_admin = self.create_action(
            name="kill_by_id_admin",
            label="Kill single process (as Admin)",
            short_desc="Kill by PID with elevated rights (taskkill /F /PID <pid>)"
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
            short_desc="Kills a processes",
            target="kill",
            args_hint=kp.ItemArgsHint.REQUIRED,
            hit_hint=kp.ItemHitHint.KEEPALL
        )

        catalog.append(killcmd)

        self.set_catalog(catalog)

    def _get_icon(self, source):
        if source in self._icons:
            return self._icons[source]
        else:
            try:
                icon = kp.load_icon("@" + source + ",0")
                self._icons[source] = icon
            except ValueError:
                self.dbg("Icon loading failed :( %s" % source)
                icon = None
            return icon

    def on_activated(self):
        """
            Creates the list of running processes, when the Keypirinha Box is
            triggered
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
                        short_desc = info["CommandLine"]
                    elif "ExecutablePath" in info and info["ExecutablePath"] != "":
                        short_desc = info["ExecutablePath"]
                    elif "Name" in info:
                        short_desc = info["Name"]

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

        self.dbg("%d running processes found" % len(self._processes))
        self.dbg("%d icons loaded" % len(self._icons))
        # self.dbg(self._icons)
        # for prc in self._processes:
        #     self.dbg(prc.raw_args())

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
            Executes the taskkill with the selected item and action
        """

        args = ["taskkill", "/F"]
        # get default action
        if action is None:
            for act in self._actions:
                if act.name() == self._default_action:
                    action = act

        # add parameters according to action
        if "kill_by_name" in action.name():
            args.append("/IM")
            # process name
            args.append(item.raw_args().split("|")[0])
        elif "kill_by_id" in action.name():
            args.append("/PID")
            # process id
            args.append(item.raw_args().split("|")[1])

        self.dbg("Calling: %s" % args)

        verb = ""
        if "_admin" in action.name():
            verb = "runas"

        # show no window when executing
        kpu.shell_execute(args[0], args[1:], verb=verb, show=subprocess.SW_HIDE)
