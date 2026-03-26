from modules.utilities.vbs_runner import VBSRunner

runner = VBSRunner()

result = runner.run("create_contract.vbs", "ZQC", "165s", "2000")

print(result.returncode)
print(result.stdout)
print(result.stderr)