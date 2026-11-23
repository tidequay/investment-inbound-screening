# Investment inbound screening

Codex reviews each application, pitch deck and public company profile. A Python helper reads and exports Excel. The only Python dependency is openpyxl.

## Install in VS Code

1. Open this skill folder in VS Code. Keep its files together when sharing it.
2. Install Python 3 if needed. On Windows, select **Add Python to PATH** in the installer if offered, then reopen VS Code.
3. Choose **Terminal > New Terminal**. Make sure the terminal is in the folder containing this README. Run:

   ```powershell
   python -m pip install -r requirements.txt
   ```

   If Windows recognizes `py` instead of `python`, use `py` in these commands.
4. Open Codex in VS Code. To run the skill from this folder, attach or reference `SKILL.md` in your request. For reuse by name across projects, copy this entire `investment-inbound-screening` skill folder to `%USERPROFILE%\.agents\skills\investment-inbound-screening` and reopen Codex. Invoke it as `$investment-inbound-screening`.

## Screen applications

1. Put your `.xlsx` export and any local decks in the workspace. Use one application per row with the column names in row 1. If applications are on another tab, include that tab's name in your request.
2. Ask Codex:

   > Follow SKILL.md to screen inbound.xlsx. Review the decks and public company links. Save a new workbook as inbound-screened.xlsx.

3. Open the new workbook and scroll right for the four screening columns. To change the thesis, edit `thesis.json` in the copy of the skill you use and run again with a new output filename.

Decks and profiles need to be accessible to Codex. Include a local copy if a deck requires access you cannot share through its link. The script reads one application tab per run and keeps other tabs in the exported copy. The working JSON files can contain private application data, so keep them with the workbook.

## Try and test it

Open `examples/inbound.xlsx` and `examples/screened.xlsx` to see the four recommendation types. The companies and facts are fictional. The example output uses the example form answers.

To try the full workflow, ask Codex to follow `SKILL.md` and screen `examples/inbound.xlsx` as fictional practice data without public lookup, saving it as `practice-screened.xlsx`.

Run the basic tests in the VS Code terminal:

```powershell
python -m unittest discover -s tests -v
```

The tests check Excel import and export, preserved links and row alignment, and rejection of missing, duplicate, or invalid results. A passing run ends with `OK`.
