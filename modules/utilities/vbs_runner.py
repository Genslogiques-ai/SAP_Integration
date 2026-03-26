import subprocess
import os

class VBSRunner:

    def __init__(self):
        # dynamically resolve script folder
        self.base_path = os.path.join(os.getcwd(),"script")

    def run(self, script_name, *args):
        script_path = os.path.join(self.base_path, script_name)

        if not os.path.exists(script_path):
            raise FileNotFoundError(f"VBS script not found: {script_path}")

        command = ["cscript", "//nologo", script_path]

        # attach arguments
        for arg in args:
            command.append(str(arg))

        result = subprocess.run(
            command,
            capture_output=True,
            text=True
        )

        return result