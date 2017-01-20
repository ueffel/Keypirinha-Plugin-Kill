import keypirinha as kp
import keypirinha_util as kpu
import os
from .lib import simpleprocess

SW_HIDE = 0

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
            short_desc="Kills all processes by that name with elevated rights (taskkill /F /IM <exe>)"
        )
        self._actions.append(kill_by_name_admin)

        kill_by_id_admin = self.create_action(
            name="kill_by_id_admin",
            label="Kill by PID (as Admin)",
            short_desc="Kills single process by its process id with elevated rights (taskkill /F /PID <pid>)"
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
            Uses the windows api
        """
        pids = simpleprocess.getpids()
        for (pid, name) in pids:
            # Don't care about the system processes, they're not killable anyway
            if pid == 0 or pid == 4:
                continue

            proc = simpleprocess.Process(pid)
            try:
                item = self.create_item(
                    category=kp.ItemCategory.KEYWORD,
                    label=name,
                    short_desc="(pid: {:5}) {}".format(pid, proc.cmdline()),
                    target="{}|{}".format(os.path.split(proc.exe())[1], str(pid)),
                    icon_handle=self._get_icon(proc.exe()),
                    args_hint=kp.ItemArgsHint.REQUIRED,
                    hit_hint=kp.ItemHitHint.IGNORE
                )
                self._processes.append(item)
                # self.dbg("Process added: {:13} {:5d} -> {}".format("normal", pid, proc.cmdline()))
            except simpleprocess.AccessDenied:
                item = self.create_item(
                    category=kp.ItemCategory.KEYWORD,
                    label=name,
                    short_desc="(pid: {:5}) Access Denied (probably only killable as Admin or not at all)".format(pid),
                    target="{}|{}".format(name, str(pid)),
                    icon_handle=None,
                    args_hint=kp.ItemArgsHint.REQUIRED,
                    hit_hint=kp.ItemHitHint.IGNORE
                )
                self._processes.append(item)
                # self.dbg("Process added: {:13} {:5d} -> {}".format("access_denied", pid, name))

        self.dbg("{:d} running processes found".format(len(self._processes)))

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
                    try:
                        self.dbg("Killing process with id: {} and name: {}".format(pid, pname))
                        proc = simpleprocess.Process(int(pid))
                        proc.kill()
                    except simpleprocess.AccessDenied:
                        self.warn("Access Denied on process '{}' (pid: {})".format(pname, pid))

        elif "kill_by_id" in action_name:
            # kill process with that pid
            try:
                self.dbg("Killing process with id: {} and name: {}".format(target_pid, target_name))
                proc = simpleprocess.Process(int(target_pid))
                proc.kill()
            except simpleprocess.AccessDenied:
                self.warn("Access Denied on process '{}' (pid: {})".format(target_name, target_pid))

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
        kpu.shell_execute(args[0], args[1:], verb="runas", show=SW_HIDE)
