# XLSX Consolidate

## Description

Consolidate just the data from several XLSX files with the exact same column headers and col data types into one XLSX combined file.

Put a copy of all your \*.xlsx files to be consolidated into one directory and call this tool to get them consolidated into one new file that gets dropped off into your home directory.

This tool simply grabs the data in each file and jambs it together but only keeps one header row. It does zero data verification, manipulation, deduplication, etc.

## Installation Dependencies

You must have `git` and `uv` installed to deploy this app to your Windows system.

See how to do that here: https://github.com/cwkingjr/windows-install-gitbash-and-uv

If you are short on Bash/Terminal Knowledge pick some up cheap here: https://github.com/cwkingjr/unix-command-intro-for-windows-folks

### If you already have UV Installed

Update it to the latest version:

```bash
uv self update
```

## Common Installation Extras

If you want to use some command line interface (CLI) tools, but want to be able to use Windows Explorer to drag a file/folder (path) and drop it on an icon to run it with your selection, you can build a Windows Batch file for the CLI app to do that. Learn how here: https://github.com/cwkingjr/windows-drag-to-app-with-args

## XLSX Consolidate Installation, Update, and Removal

### Install

Using the Git Bash Terminal run this command at the shell prompt:

```bash
uv tool install https://github.com/cwkingjr/xlsx-consolidate.git
```

### Upgrade

This is typically only needed when the developer fixes bugs, adds features, or updates the program's dependencies. In any case, it doesn't hurt anything to run it just in case.

```bash
uv tool upgrade xlsx-consolidate
```

### Uninstall

```bash
uv tool uninstall xlsx-consolidate
```

If you made/deployed Windows Batch files to your desktop, you'll need to manually delete those.

### List UV-Managed Tools

```bash
uv tool list
```

## Tools Included

- xlsx_consolidate # Linux/Mac
- xlsx_consolidate.exe # Windows

### Invoke XLSX Consolidate

To run the tools, go to the Git Bash Terminal and run:

```bash
xlsx_consolidate.exe  # with no argument, prints usage help
xlsx_consolidate.exe -h or --help  # prints the full help info
xlsx_consolidate.exe -d <xlsx-directory-path> # replace arrow tags & contents with the directory path
xlsx_consolidate.exe --directory <xlsx-directory-path>  # -d/--directory do the same thing
xlsx_consolidate.exe -d ~/Documents/xlsx_reports/  # ~ is replaced by user's home directory in bash
xlsx_consolidate.exe --directory ~/Documents/xlsx_reports/
```

#### Example CLI Output When Finished

```bash
'Consolidated XLSX file written to /Users/<username>/Documents/xlsx_consolidated_20250814_004541.xlsx'
Note: /Users/<username> is a home directory on Mac
```

### Example Batch File

If your system installs this application in the same area as mine does, you should be able to download this file https://github.com/cwkingjr/xlsx-consolidate/blob/main/xlsx_consolidate.bat onto your desktop and use Windows (File) Explorer to drag a CSV file path to it and drop it on the icon to invoke the app. If it doesn't work, see the installation extras section above to reconfigure it for your installation path.
