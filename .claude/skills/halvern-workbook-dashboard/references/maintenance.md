# Maintaining the skill

- Change a mapping, threshold or rule text in `scripts/config.py` only. The dashboard rules table and the script both read it.
- A new type of check needs a change to `scripts/process.py` as well, and only when Jeremie asks.
- After any change, rerun on the last workbook and confirm the output is as Jeremie expects before relying on it.
- To compare two versions, run the script on both and compare the printed output and issue lists. Do not compare the workbooks by eye.
- Q6_TREATMENT in the config records that Jeremie confirmed the Portfolio tab as correct. Change it if that decision changes.
