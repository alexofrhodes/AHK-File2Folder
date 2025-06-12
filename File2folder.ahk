#Requires AutoHotkey v2.0
#SingleInstance Force

iniFile := A_ScriptDir "\config.ini"
hotkeySection := "Settings"
hotkeyKey := "Hotkey"
defaultHotkey := "F1"

; Read or create config.ini
if !FileExist(iniFile) {
    IniWrite defaultHotkey, iniFile, hotkeySection, hotkeyKey
    hk := defaultHotkey
} else {
    hk := IniRead(iniFile, hotkeySection, hotkeyKey, defaultHotkey)
}

; Register dynamic hotkey
Hotkey hk, MoveEachFileToItsOwnFolder

MoveEachFileToItsOwnFolder(*) {
    selectedFiles := GetSelectedFilesInExplorer()
    if !selectedFiles.Length {
        MsgBox "No files selected."
        return
    }

    for filePath in selectedFiles {
        if !FileExist(filePath)
            continue

        SplitPath filePath, &fileName, &dir, &ext
        baseName := StrReplace(fileName, "." . ext)

        targetDir := dir "\" baseName

        if DirExist(targetDir) {
            ts := FormatTime(, "yyyy-MM-dd_HH-mm-ss")
            targetDir := dir "\" baseName "_" ts
        }

        DirCreate(targetDir)

        try {
            FileMove filePath, targetDir "\" fileName "." ext
        } catch Error as e {
            MsgBox "Failed to move " fileName ": " e.Message
        }
    }
}

GetSelectedFilesInExplorer() {
    shell := ComObject("Shell.Application")
    selected := []

    for window in shell.Windows {
        try {
            if InStr(window.FullName, "explorer.exe") = 0
                continue

            items := window.Document.SelectedItems()
            for item in items
                selected.Push(item.Path)
        } catch
            continue
    }

    return selected
}
