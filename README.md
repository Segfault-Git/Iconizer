# 🎨  Iconizer - Icon Extraction & Folder Skinner 🚀

<p align="center">
  <img src="images/main.png" alt="Iconizer Main Icon"/>
</p>
Iconizer is a lightweight PowerShell tool designed to extract icons from .exe files and assigning custom icons files to folders with zero installation, configuration, or leftover traces.

The core idea is simplicity and portability: just run the script from anywhere and get immediate results. No dependencies, no setup — just clean PowerShell with c# seasoning and native Windows functionality.

The script consists of two parts. The first part, **`pull`**, involves extracting icons from executable files (`.exe`) and saving them in different formats. The second part, **`apply`**, customizes folder icons by locating `.exe` or `.ico` files and creating a `desktop.ini` file with the required settings to display an icon on a folder, making both tools useful for developers and users who need efficient icon management solutions.

## 📋 Table of Contents

- [✨ Features](#-features)
- [💡 Quick Start](#-quick-start)
- [🔧 Icon Extraction (`pull`)](#-icon-extraction-pull)
- [🔧 Folder Customization (`apply`)](#-folder-customization-apply)
- [⚡ Important Notes](#-important-notes)
- [🔍 Troubleshooting](#-troubleshooting)
- [📄 License](#-license)
- [🤝 Contributing](#-contributing)

## ✨ Features

### 🛠️  Icon Extraction (`pull`)

- Extract icons from executable files (.exe) in various formats
- Support for ICO and PNG output formats
- Extract specific icon groups by index or all available icons
- High-quality extraction with support for large icons (up to 256x256)
- Automatic detection of PNG-embedded icons
- Comprehensive logging and error handling
- GUI file selection dialog

DEMO

<img src="./images/pull.gif" width="637" height="447" />


### 🖼️  Folder Customization (`apply`)

- Automatically apply icons to folders using found .exe or .ico files
- Intelligent icon matching based on folder names
- Priority system for icon selection
- Bulk processing with recursive folder scanning
- Custom filtering to exclude specific folders
- Desktop.ini file management
- Icon removal functionality with Explorer restart

DEMO

<img src="./images/apply.gif" width="637" height="447" />

## 💡 Quick Start

>[!NOTE]
> Direct link for security-conscious users:
<https://raw.githubusercontent.com/JustDj-Git/Iconizer/main/Iconizer.ps1>

1. Open PowerShell (not CMD). Right-click on the Windows start menu and find PowerShell (or Terminal), or press `Win + S` and type Powershell.
2. Copy and paste the code below and press enter for invoking **`pull`** function only

```powershell
irm icon.scripts.wiki | iex; pull
```

or for invoking **`apply`** function only

```powershell
irm icon.scripts.wiki | iex; apply
```

or for both

```powershell
irm icon.scripts.wiki | iex; pull; apply
```

## 🔧 Icon Extraction (`pull`)

Extract icons from executable files with various options and formats.

 📝 The script accepts the following parameters and switches:

| Parameter      | Alias  | Type     | Description                           | Examples                         |
|----------------|--------|----------|---------------------------------------|----------------------------------|
| `-directory`   | `-d`   | string[] | Paths to .exe or several directories  | `-d "FIRST_PATH", "SECOND_PATH"` |
| `-index`       | `-i`   | int      | Icon group index to extract           | `-i 1`                           |
| `-search_depth`| `-sd`  | int      | Subdirectory recursion depth          | `-sd 2`                          |
| `-png`         |        | switch   | Extract largest icon as PNG           |                                  |
| `-info`        |        | switch   | Show icons info (no extraction)       |                                  |
| `-all`         | `-a`   | switch   | Extract all icons                     |                                  |
| `-file`        | `-f`   | switch   | Opens a file selection dialog         |                                  |
| `-log`         | `-l`   | string   | Enable logging to specified file path | `-l 'C:\'`                       |
| `-pause`       |        | switch   | Pause script in the end               |                                  |
| `-verbose`     |        | switch   | Show all verbose information          |                                  |

>[!TIP]
> If **`directory`** is not specified, a system FileDialog will open.
> If **`png`** is not specified, `ico` will be used as the default format.
> If **`index`** is not specified, index `0` will be used by default.

### ⚙️ Usage Examples (`pull`)

#### Basic Icon Extraction

```powershell
# Extract from single executable
irm icon.scripts.wiki | iex; pull -d 'C:\Program Files\MyApp\app.exe'

# Extract from multiple paths
irm icon.scripts.wiki | iex; pull -d 'C:\Apps\', 'D:\Games\game.exe'
```

#### Advanced Extraction Options

```powershell
# Extract all icons as PNG format with logging
irm icon.scripts.wiki | iex; pull -d 'C:\Apps\' -png -all -log 'C:\'

# Extract specific icon group with recursive search
irm icon.scripts.wiki | iex; pull -d 'C:\Programs\' -index 2 -depth 3

# Get icon information without extraction
irm icon.scripts.wiki | iex; pull -d 'C:\app.exe' -info
```

## 🎯 Folder Customization (`apply`)

📝 The script accepts the following parameters and switches:

| Parameter       | Alias  | Type      | Description                            | Examples                                  |
|-----------------|--------|-----------|----------------------------------------|-------------------------------------------|
| `-directory`    | `-d`   | string[]  | Folder paths to process                | -d `'FIRST_PATH', 'SECOND_PATH'`          |
| `-priority`     | `-p`   | string    | Icon selection priority (`ico`, `exe`) | -p `'exe'`                                |
| `-filter`       | `-f`   | string[]  | Folder names to exclude                | -f `'FIRST_NAME', 'SECOND_NAME'`          |
| `-rules`        | `-r`   | hashtable | Custom icon applying rules             | -r `@{"X" = "N.exe"; "Y" = "Z.exe"}`      |
| `-noforce`      | `-nf`  | switch    | Skip folders with existing desktop.ini |                                           |
| `-search_depth` | `-sd`  | int       | Icon file search depth                 | `-sd 2`                                   |
| `-apply_depth`  | `-ad`  | int       | Folder processing depth                | `-ad 2`                                   |
| `-remove`       | `-rm`  | switch    | Remove folder icons                    |                                           |
| `-log`          | `-l`   | string    | Enable logging to specified path       | `-l 'C:\'`                                |
| `-pause`        |        | switch    | Pause script in the end                |                                           |
| `-verbose`      |        | switch    | Show all verbose information           |                                           |

>[!TIP]
> If **`directory`** is not specified, a system FileDialog will open.
> If multiple .ico or .exe files are found and the script does not find any filename matching the folder name, it will select the first suitable file from the list. You can rename the .ico file to match the folder name (or similar one) to prevent this from happening.

### ⚙️ Examples with `apply`

```powershell
# Apply icons to single folder
irm icon.scripts.wiki | iex; apply -d 'D:\Programs'

# Apply icons to multiple folders
irm icon.scripts.wiki | iex; apply -d 'C:\Games', 'D:\Apps'
```

#### Advanced Options

```powershell
# Prioritize .ico files over .exe files
irm icon.scripts.wiki | iex; apply -d 'D:\Programs' -p ico

# Apply to single folder without recursion
irm icon.scripts.wiki | iex; apply -d 'C:\MyApp' -s

# Exclude specific folders from processing
irm icon.scripts.wiki | iex; apply -d 'D:\Programs' -f 'Backup', 'Temp'
```

#### Custom Dependencies

```powershell
# Define custom icon applying rules
$rules = @{
    "Visual Studio Code" = "Code.exe"
    "Adobe Photoshop" = "Photoshop.exe"
    "Steam" = "steam.exe"
}
irm icon.scripts.wiki | iex; apply -d 'D:\Programs' -r $rules
```

#### Icon Removal

```powershell
# Remove all folder icons and restart Explorer
irm icon.scripts.wiki | iex; apply -d 'D:\Programs' -rm
```

## 🔍 Troubleshooting

### Permission Errors

```powershell
# Run PowerShell as Administrator for system folders
# Or exclude protected folders with filter:
irm icon.scripts.wiki | iex; apply -d 'C:\' -filter 'Windows', 'System32', 'Program Files'
```

### Icon Extraction Failures

```powershell
# Use info mode to check available icons first
irm icon.scripts.wiki | iex; pull -d 'problematic.exe' -info

# Try different icon groups
irm icon.scripts.wiki | iex; pull -d 'app.exe' -index 2

# Try to extract all icons
irm icon.scripts.wiki | iex; pull -d 'app.exe' -all
```

## ⚡ Important Notes

### System Requirements

- Windows PowerShell 5.1 or PowerShell Core 6+
- .NET Framework for image processing
- Administrative privileges for some system folders

### File Safety

- Always test on a small folder set first
- Use `-noforce` flag to preserve existing customizations

### Performance Considerations

- Large directory trees may take significant time to process
- Use appropriate `-depth` values to limit recursion
- Consider using `-info` mode for initial assessment

## 📄 License

This project is provided as-is for educational and personal use. Please respect software licenses when extracting icons from commercial applications.

## 🤝 Contributing

Feel free to submit issues, feature requests, or improvements to make this tool even better for the community!
