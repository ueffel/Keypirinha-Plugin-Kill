import keypirinha as kp
import keypirinha_util as kpu
import subprocess

class Kill(kp.Plugin):
    """
        Plugin that lists running processes with name and commandline if
        available and kills the select process with
        'taskkill.exe /F /IM <selected process exe>'
    """

    def __init__(self):
        super().__init__()
        self._processes = []
        # self._debug = True

    def on_catalog(self):
        """
            Adds the kill command to the catalog
        """
        catalog = []

        killcmd = self.create_item(
            category=kp.ItemCategory.KEYWORD,
            label="Kill Process:",
            short_desc="Kills a process with taskkill.exe",
            target="kill",
            args_hint=kp.ItemArgsHint.REQUIRED,
            hit_hint=kp.ItemHitHint.KEEPALL
        )

        catalog.append(killcmd)

        self.set_catalog(catalog)

    def on_activated(self):
        """
            Creates the list of running processes, when the Keypirinha Box is
            triggered
        """
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output, err = subprocess.Popen(["wmic",
                                        "process",
                                        "get",
                                        "ProcessId,Caption,Description,",
                                        "Name,ExecutablePath,CommandLine",
                                        "/FORMAT:CSV"],
                                       stdout=subprocess.PIPE,
                                       startupinfo=startupinfo).communicate()

        if err:
            self.err(err)

        initial_item = self.create_item(
            category=kp.ItemCategory.KEYWORD,
            label="Kill Process:",
            short_desc="Kills a process by name with taskkill.exe",
            target="kill",
            args_hint=kp.ItemArgsHint.REQUIRED,
            hit_hint=kp.ItemHitHint.KEEPALL
        )

        indexes = {}
        header_read = False

        for line in output.splitlines():
            if line.strip() == b"":
                continue
            if not header_read:
                header = line.split(b",")
                for idx, col in enumerate(header):
                    if col == b"Node":
                        continue
                    else:
                        indexes[col] = idx
                header_read = True
            else:
                if line[indexes[b"Caption"]] == b"System Idle Process" \
                    or line[indexes[b"Caption"]] == b"System":
                    continue

                cols = line.split(b",")
                item = initial_item.clone()
                item.set_args(cols[indexes[b"Name"]].decode(), cols[indexes[b"Caption"]].decode())
                if cols[indexes[b"CommandLine"]]:
                    item.set_short_desc(cols[indexes[b"CommandLine"]].decode())
                elif cols[indexes[b"ExecutablePath"]]:
                    item.set_short_desc(cols[indexes[b"ExecutablePath"]].decode())
                elif cols[indexes[b"Name"]]:
                    item.set_short_desc(cols[indexes[b"Name"]].decode())
                self._processes.append(item)

        self.dbg("%d running processes found" % len(self._processes))

    def on_deactivated(self):
        """
            Emptys the process list, when Keypirinha Box is closed
        """
        self._processes = []

    def on_suggest(self, user_input, initial_item=None, current_item=None):
        """
            Sets the list of running processes as suggestions
        """
        if not initial_item:
            return

        self.set_suggestions(self._processes, kp.Match.FUZZY, kp.Sort.SCORE_DESC)

    def on_execute(self, item, action):
        """
            Executes the taskkill with the selected item
        """
        self.dbg("killing process '%s'" % item.raw_args())

        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.Popen(["taskkill", "/F", "/IM", item.raw_args()],
                         startupinfo=startupinfo).communicate()
