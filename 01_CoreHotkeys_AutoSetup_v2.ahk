#Requires AutoHotkey v2.0
; ============================================================
; Falcon Electric — Core Hotkeys, Snippets, & INI Shortcuts
; AutoHotkey v2.0.19  |  Build v3.0-au (quotes fixed + updater)
; ============================================================

#SingleInstance Force
#Warn
SetTitleMatchMode 2

; ----------------------- Globals -----------------------
INI := A_ScriptDir "\FalconAHK.ini"
PathsByHotkey := Map()                 ; hotkey -> target path
AppName  := "Falcon AHK"
AppVer   := "3.0.0"                    ; current script version
AppBuild := "2025-08-11"

; Defaults (ASCII only)
default_JobRoot     := "C:\Falcon\Jobs"
default_PermitsRoot := "C:\Falcon\Permits"
default_ServiceRoot := "C:\Falcon\Service Work"
default_ArrivalWin  := "7:00-7:30 AM"
default_ArrivalWin2 := "10:00-11:00 AM"
default_TBExe       := "thunderbird.exe"
default_TagSlot     := "1"

; Update defaults
default_AutoUpdate := "0"              ; 0=off, 1=on
default_UpdateURL  := ""               ; e.g. https://raw.githubusercontent.com/your/repo/main/falcon_update.json

; Live settings
JobRoot := "", PermitsRoot := "", ServiceRoot := ""
ArrivalWin := "", ArrivalWin2 := "", TBExe := "", TBExeProc := "", TagSlot := ""
AutoUpdate := "", UpdateURL := ""

; -------------------- Tray menu ------------------------
TraySetIcon("shell32.dll", 44)
try A_TrayMenu.Delete()
A_TrayMenu.Add("Open Setup", (*) => OpenSetupGUI())
A_TrayMenu.Add("Manage Shortcuts", (*) => ShowShortcutsManager())
A_TrayMenu.Add("Reload INI Shortcuts", (*) => (InitShortcuts(), TrayTip(AppName,"Shortcuts reloaded",1)))
A_TrayMenu.Add("Open INI", (*) => Run(INI))
A_TrayMenu.Add()
A_TrayMenu.Add("Check for Updates", (*) => CheckForUpdate(false))
A_TrayMenu.Add("Toggle Auto-Update", (*) => ToggleAutoUpdate())
A_TrayMenu.Add()
A_TrayMenu.AddStandard()
TrayTip(AppName,"Loaded",1)

; ------------------ Load + init ------------------------
LoadOrSetup()
InitShortcuts()
if (AutoUpdate = "1")
    SetTimer(() => CheckForUpdate(true), -500) ; one-shot silent check shortly after startup

; ===================== HOTSTRINGS ======================
::;arr::  { SendText(Format("Our arrival window is {}. We'll call if anything changes.", ArrivalWin)) }
::;arr2:: { SendText(Format("Our arrival window is {}. We'll call if anything changes.", ArrivalWin2)) }

::;conf:: {
    r := InputBox("Enter the date (e.g., 08/12/2025):", "Confirm - Date")
    if (r.Result = "Cancel")
        return
    choice := MsgBox(Format("Use 1st stop ({})?`nYes = 1st, No = 2nd, Cancel = abort.", ArrivalWin), "Arrival Window", "YesNoCancel Iconi")
    if (choice = "Cancel")
        return
    apptTime := (choice = "Yes") ? ArrivalWin : ArrivalWin2
    SendText(Format("Just confirming we are scheduled for {} at {}. Please reply if this time no longer works.", r.Value, apptTime))
}
::;eta::  { SendText("Our tech should be there in approximately [X] minutes.") }
::;fup::  { SendText("Following up to see if you had a chance to review my last message.") }
::;sig::  {
    sig := "Eric Eder{Enter}"
        . "Office Manager{Enter}"
        . "Falcon Electric Inc.{Enter}"
        . "Office: (847) 883-9913 | Cell: (847) 812-8305{Enter}"
        . "Email: falconelectriceric@comcast.net{Enter}"
        . "Website: www.falconelectric.net"
    Send(sig)
}

; =================== UTILITY HOTKEYS ===================
^!F7::SendText(ArrivalWin)        ; paste 1st window
^!F8::SendText(ArrivalWin2)       ; paste 2nd window

^!d:: {
    today := FormatTime(, "MM/dd/yyyy")
    SendText(today)
}
^!t:: {
    now := FormatTime(, "h:mm tt")
    SendText(now)
}
^!v:: {
    clipSaved := ClipboardAll()
    ClipWait(0.3)
    A_Clipboard := A_Clipboard
    Send("^v")
    Sleep 120
    A_Clipboard := clipSaved
}
^!c:: {
    clipSaved := ClipboardAll()
    ClipWait(0.3)
    plain := A_Clipboard
    plain := StrReplace(plain, "`r`n", "`n")
    plain := StrReplace(plain, "`r", "`n")
    plain := RegExReplace(plain, "m)^[ \t]+", "")
    plain := RegExReplace(plain, "m)[ \t]+$", "")
    A_Clipboard := plain
    SoundBeep(1000)
    Sleep 150
    A_Clipboard := clipSaved
}
^!y:: {
    stamp := FormatTime(, "yyyy-MM-dd")
    SendText(stamp " - ")
}
^!F12::OpenSetupGUI()

; ================== FOLDER JUMP HOTKEYS =================
^!j:: Run(JobRoot)
^!p:: Run(PermitsRoot)
^!s:: Run(ServiceRoot)

; ============== THUNDERBIRD-SPECIFIC HOTKEYS ============
^!1::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("^n")
}
^!2::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("^r")
}
^!3::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("+^r")
}
^!4::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("^f")
}
^!5::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("^k")
}
^!a::
{
    if WinActive("ahk_exe " TBExeProc)
        Send("a")
}
^!g::
{
    if WinActive("ahk_exe " TBExeProc) && TagSlot != ""
        Send(TagSlot)
}

; ======= INI Shortcuts: open INI / reload & manager =======
^!+I::Run(INI)
^!+R::(InitShortcuts(), MsgBox("Shortcuts reloaded.","Reloaded","OK Iconi"))
^!+M::ShowShortcutsManager()

; ==========================================================
;                     FUNCTIONS (v2)
; ==========================================================
NormalizeTBExe() {
    global TBExe, TBExeProc
    if InStr(TBExe, "\") || InStr(TBExe, "/") {
        SplitPath(TBExe, &file)
        TBExeProc := file
    } else {
        TBExeProc := TBExe
    }
}

LoadOrSetup() {
    global
    if !FileExist(INI) {
        JobRoot     := default_JobRoot
        PermitsRoot := default_PermitsRoot
        ServiceRoot := default_ServiceRoot
        ArrivalWin  := default_ArrivalWin
        ArrivalWin2 := default_ArrivalWin2
        TBExe       := default_TBExe
        TagSlot     := default_TagSlot
        AutoUpdate  := default_AutoUpdate
        UpdateURL   := default_UpdateURL
        SaveSettingsToIni()
        NormalizeTBExe()
        OpenSetupGUI()
        return
    }
    JobRoot     := IniRead(INI, "Paths", "JobRoot",     default_JobRoot)
    PermitsRoot := IniRead(INI, "Paths", "PermitsRoot", default_PermitsRoot)
    ServiceRoot := IniRead(INI, "Paths", "ServiceRoot", default_ServiceRoot)
    ArrivalWin  := IniRead(INI, "Email", "ArrivalWindow",  default_ArrivalWin)
    ArrivalWin2 := IniRead(INI, "Email", "ArrivalWindow2", default_ArrivalWin2)
    TBExe       := IniRead(INI, "Email", "ThunderbirdExe", default_TBExe)
    TagSlot     := IniRead(INI, "Email", "TagSlot",        default_TagSlot)
    AutoUpdate  := IniRead(INI, "Update", "AutoUpdate",    default_AutoUpdate)
    UpdateURL   := IniRead(INI, "Update", "UpdateURL",     default_UpdateURL)
    NormalizeTBExe()
}

SaveSettingsToIni() {
    global
    IniWrite(JobRoot,     INI, "Paths", "JobRoot")
    IniWrite(PermitsRoot, INI, "Paths", "PermitsRoot")
    IniWrite(ServiceRoot, INI, "Paths", "ServiceRoot")
    IniWrite(ArrivalWin,  INI, "Email", "ArrivalWindow")
    IniWrite(ArrivalWin2, INI, "Email", "ArrivalWindow2")
    IniWrite(TBExe,       INI, "Email", "ThunderbirdExe")
    IniWrite(TagSlot,     INI, "Email", "TagSlot")
    IniWrite(AutoUpdate,  INI, "Update", "AutoUpdate")
    IniWrite(UpdateURL,   INI, "Update", "UpdateURL")
}

; --------------------- Setup GUI -------------------------
OpenSetupGUI() {
    global
    static g1 := 0
    if g1 {
        g1.Show()
        return
    }
    g1 := Gui("+AlwaysOnTop", "Falcon AHK - Setup")
    g1.MarginX := 12, g1.MarginY := 12
    g1.SetFont("s9", "Segoe UI")
    g1.Add("Text",, "Configure Falcon Electric hotkeys, folders, and email behavior:")

    g1.Add("GroupBox", "x12 y40 w520 h150", "Folders")
    g1.Add("Text",  "x24  y68  w90", "Job Root:")
    eJob := g1.Add("Edit",  "x116 y64  w320 vGUI_JobRoot", JobRoot)
    btn := g1.Add("Button","x444 y64  w72", "Browse")
    btn.OnEvent("Click", (*) => (tmp := DirSelect("", 1, "Select Job Root"), tmp ? eJob.Value := tmp : ""))

    g1.Add("Text",  "x24  y100 w90", "Permits Root:")
    ePerm := g1.Add("Edit",  "x116 y96  w320 vGUI_PermitsRoot", PermitsRoot)
    btn := g1.Add("Button","x444 y96  w72", "Browse")
    btn.OnEvent("Click", (*) => (tmp := DirSelect("", 1, "Select Permits Root"), tmp ? ePerm.Value := tmp : ""))

    g1.Add("Text",  "x24  y132 w90", "Service Root:")
    eSvc := g1.Add("Edit",  "x116 y128 w320 vGUI_ServiceRoot", ServiceRoot)
    btn := g1.Add("Button","x444 y128 w72", "Browse")
    btn.OnEvent("Click", (*) => (tmp := DirSelect("", 1, "Select Service Root"), tmp ? eSvc.Value := tmp : ""))

    g1.Add("GroupBox", "x12 y200 w520 h190", "Email / Update Settings")
    g1.Add("Text", "x24  y228 w210", "Arrival Window (e.g., 7:00-7:30 AM):")
    eArr1 := g1.Add("Edit", "x238 y224 w110 vGUI_ArrivalWin", ArrivalWin)
    g1.Add("Text", "x360 y228 w90", "Thunderbird EXE:")
    eTB := g1.Add("Edit", "x454 y224 w78 vGUI_TBExe", TBExe)

    labelTag := 'Thunderbird Tag Slot for "To-Schedule" (1-9 or blank):'
    g1.Add("Text", "x24  y262 w300", labelTag)
    eTag := g1.Add("Edit", "x324 y258 w60 vGUI_TagSlot", TagSlot)

    g1.Add("Text", "x24  y296 w260", "Second Arrival Window (e.g., 10:00-11:00 AM):")
    eArr2 := g1.Add("Edit", "x288 y292 w110 vGUI_ArrivalWin2", ArrivalWin2)

    g1.Add("Text", "x24  y328 w210", "Update URL (manifest or JSON):")
    eURL := g1.Add("Edit", "x238 y324 w294 vGUI_UpdateURL", UpdateURL)
    cbAU := g1.Add("CheckBox", "x24 y352 vGUI_AutoUpdate", "Enable Auto-Update on launch")
    cbAU.Value := (AutoUpdate = "1") ? 1 : 0

    btn := g1.Add("Button", "x12  y392 w120 Default", "Save & Apply")
    btn.OnEvent("Click", (*) => SaveSettings_Click())
    btn := g1.Add("Button", "x142 y392 w120", "Test Snippet")
    btn.OnEvent("Click", TestSnippet)
    btn := g1.Add("Button", "x272 y392 w120", "Manage Shortcuts")
    btn.OnEvent("Click", (*) => ShowShortcutsManager())
    btn := g1.Add("Button", "x402 y392 w120", "Close")
    btn.OnEvent("Click", (*) => g1.Hide())

    g1.Show("w544 h430")

    SaveSettings_Click() {
        g1.Submit(false)
        JobRoot     := g1["GUI_JobRoot"].Value
        PermitsRoot := g1["GUI_PermitsRoot"].Value
        ServiceRoot := g1["GUI_ServiceRoot"].Value
        ArrivalWin  := g1["GUI_ArrivalWin"].Value
        ArrivalWin2 := g1["GUI_ArrivalWin2"].Value
        TBExe       := g1["GUI_TBExe"].Value
        TagSlot     := g1["GUI_TagSlot"].Value
        UpdateURL   := g1["GUI_UpdateURL"].Value
        AutoUpdate  := g1["GUI_AutoUpdate"].Value ? "1" : "0"
        NormalizeTBExe()
        SaveSettingsToIni()
        MsgBox("Settings saved. Hotkeys are live.", "Saved", "OK Iconi")
    }
}

; ---------------- Helpers & Events ----------------------
TestSnippet(*) {
    MsgBox("Type ;arr or ;arr2 in any text field to test.", "Test Snippet", "OK Iconi")
}

ToggleAutoUpdate() {
    global AutoUpdate
    AutoUpdate := (AutoUpdate = "1") ? "0" : "1"
    IniWrite(AutoUpdate, INI, "Update", "AutoUpdate")
    MsgBox("Auto-Update is now set to: " (AutoUpdate="1" ? "ON" : "OFF"), "Auto-Update", "OK Iconi")
}

; ---------------- INI-Driven Shortcuts -------------------
InitShortcuts() {
    global
    EnsureIniShortcut(1, "Schedule Excel",                "^!+F1", "C:\Falcon\Admin\Schedule.xlsx")
    EnsureIniShortcut(2, "Call Sheet Excel",              "^!+F2", "C:\Falcon\Admin\Call Sheet.xlsx")
    EnsureIniShortcut(3, "Panel Schedule Folder",         "^!+F3", "C:\Falcon\Panel Schedules\")
    EnsureIniShortcut(4, "Job Plan Folder",               "^!+F4", "C:\Falcon\Plans\")
    EnsureIniShortcut(5, "Panel Schedule Blank 42 Space", "^!+F5", "C:\Falcon\Templates\Panel Schedule Blank 42.xlsx")
    EnsureIniShortcut(6, "Panel Schedule Blank 66 Space", "^!+F6", "C:\Falcon\Templates\Panel Schedule Blank 66.xlsx")

    for hk, _ in PathsByHotkey
        Hotkey(hk, "Off")
    PathsByHotkey := Map()

    idx := 1
    loop {
        name   := IniRead(INI, "Shortcuts", idx "_Name",   "")
        hotkey := IniRead(INI, "Shortcuts", idx "_Hotkey", "")
        path   := IniRead(INI, "Shortcuts", idx "_Path",   "")
        if (name = "" && hotkey = "" && path = "")
            break
        if (hotkey != "" && path != "") {
            PathsByHotkey[hotkey] := path
            Hotkey(hotkey, OpenShortcutTarget, "On")
        }
        idx++
    }
}

EnsureIniShortcut(idx, name, hotkey, path) {
    if (IniRead(INI, "Shortcuts", idx "_Name",   "") = "")
        IniWrite(name,   INI, "Shortcuts", idx "_Name")
    if (IniRead(INI, "Shortcuts", idx "_Hotkey", "") = "")
        IniWrite(hotkey, INI, "Shortcuts", idx "_Hotkey")
    if (IniRead(INI, "Shortcuts", idx "_Path",   "") = "")
        IniWrite(path,   INI, "Shortcuts", idx "_Path")
}

OpenShortcutTarget(*) {
    global PathsByHotkey
    hk := A_ThisHotkey
    if !PathsByHotkey.Has(hk)
        return
    target := PathsByHotkey[hk]
    if !FileExist(target) {
        MsgBox("Target not found:`n" target, "Not Found", "Icon!")
        return
    }
    Run(target)
}

; -------------------- Updater ----------------------------
CheckForUpdate(silent := true) {
    global AppVer, INI, UpdateURL
    if (UpdateURL = "") {
        if !silent
            MsgBox("No Update URL set. Add one in Setup > Update URL.", "Updater", "OK Iconi")
        return
    }
    try {
        tmpFile := A_Temp "\falcon_update.tmp"
        Download(UpdateURL, tmpFile)  ; works with JSON or INI/text
        txt := FileRead(tmpFile, "UTF-8")
        FileDelete(tmpFile)
    } catch e {
        if !silent
            MsgBox("Download failed:`n" e.Message, "Updater", "OK Icon!")
        return
    }
    ver := "", url := ""
    ParseUpdateManifest(txt, &ver, &url)
    if (ver = "" || url = "") {
        if !silent
            MsgBox("Invalid update manifest. Expect INI or JSON with version/url.", "Updater", "OK Icon!")
        return
    }
    if (VersionCompare(ver, AppVer) <= 0) {
        if !silent
            MsgBox("You're up to date (" AppVer ").", "Updater", "OK Iconi")
        return
    }
    if (MsgBox("Update " AppVer " → " ver "?`nProceed to download and replace the script?", "Update Available", "YesNo Iconi") != "Yes")
        return

    try {
        newFile := A_Temp "\falcon_new.ahk"
        Download(url, newFile)
        backup := A_ScriptFullPath ".bak"
        FileCopy(A_ScriptFullPath, backup, 1)
        FileMove(newFile, A_ScriptFullPath, 1)
        Run('"' A_AhkPath '" "' A_ScriptFullPath '"')
        ExitApp
    } catch e {
        MsgBox("Update failed:`n" e.Message, "Updater", "OK Icon!")
    }
}

ParseUpdateManifest(txt, &ver, &url) {
    ; Supports either JSON: {"version":"x.y.z","url":"https://..."}
    ; or INI-like: Version=x.y.z  URL=https://...
    ver := "", url := ""
    ; JSON quick parse (no external libs): find "version" and "url"
    if InStr(txt, "{") && InStr(txt, "}") {
        if m := RegExMatch(txt, '"version"\s*:\s*"([^"]+)"', &mm)
            ver := mm[1]
        if m := RegExMatch(txt, '"url"\s*:\s*"([^"]+)"', &mm)
            url := mm[1]
        return
    }
    ; INI-like
    for line in StrSplit(txt, "`n") {
        line := Trim(StrReplace(line, "`r"))
        if RegExMatch(line, "i)^version\s*=\s*(.+)$", &m1)
            ver := Trim(m1[1])
        else if RegExMatch(line, "i)^url\s*=\s*(.+)$", &m2)
            url := Trim(m2[1])
    }
}

VersionCompare(a, b) {
    ; returns 1 if a>b, 0 if a=b, -1 if a<b
    pa := StrSplit(a, "."), pb := StrSplit(b, ".")
    max := (pa.Length > pb.Length) ? pa.Length : pb.Length
    loop max {
        ai := (A_Index <= pa.Length) ? Integer(pa[A_Index]) : 0
        bi := (A_Index <= pb.Length) ? Integer(pb[A_Index]) : 0
        if (ai > bi)
            return 1
        if (ai < bi)
            return -1
    }
    return 0
}

; --------------- Shortcuts Manager GUI -------------------
ShowShortcutsManager() {
    global
    static g2 := 0, lv := 0
    if !g2 {
        g2 := Gui(, "Shortcuts Manager")
        g2.MarginX := 12, g2.MarginY := 12
        g2.SetFont("s9", "Segoe UI")
        g2.Add("Text",, "Manage file/folder shortcuts:")
        lv := g2.Add("ListView", "x12 y30 w520 h200", ["ID","Name","Hotkey","Path"])
        btn := g2.Add("Button", "x12  y240 w80",  "Add")
        btn.OnEvent("Click", (*) => AddShortcut())
        btn := g2.Add("Button", "x102 y240 w80",  "Edit")
        btn.OnEvent("Click", (*) => EditShortcut())
        btn := g2.Add("Button", "x192 y240 w80",  "Delete")
        btn.OnEvent("Click", (*) => DeleteShortcut())
        btn := g2.Add("Button", "x282 y240 w80",  "Reload")
        btn.OnEvent("Click", (*) => (InitShortcuts(), RefreshList()))
        btn := g2.Add("Button", "x372 y240 w80",  "Close")
        btn.OnEvent("Click", (*) => g2.Hide())
        g2.OnEvent("Close", (*) => g2.Hide())
    }
    RefreshList()
    g2.Show("w550 h280")

    RefreshList() {
        lv.Delete()
        idx := 1
        loop {
            name   := IniRead(INI, "Shortcuts", idx "_Name",   "")
            hotkey := IniRead(INI, "Shortcuts", idx "_Hotkey", "")
            path   := IniRead(INI, "Shortcuts", idx "_Path",   "")
            if (name = "" && hotkey = "" && path = "")
                break
            lv.Add("", idx, name, hotkey, path)
            idx++
        }
        lv.ModifyCol(1, 35), lv.ModifyCol(2, 130), lv.ModifyCol(3, 120), lv.ModifyCol(4, 330)
    }

    AddShortcut()  => ShowAddEditGui()
    EditShortcut() {
        row := lv.GetNext()
        if !row {
            MsgBox("Please select a shortcut to edit.", "Error", "Icon!")
            return
        }
        id := lv.GetText(row, 1)
        ShowAddEditGui(id)
    }
    DeleteShortcut() {
        row := lv.GetNext()
        if !row {
            MsgBox("Please select a shortcut to delete.", "Error", "Icon!")
            return
        }
        id := lv.GetText(row, 1)
        if (MsgBox("Delete selected shortcut?", "Confirm", "YesNo Icon?") = "Yes") {
            IniDelete(INI, "Shortcuts", id "_Name")
            IniDelete(INI, "Shortcuts", id "_Hotkey")
            IniDelete(INI, "Shortcuts", id "_Path")
            InitShortcuts()
            RefreshList()
        }
    }

    ShowAddEditGui(editId := "") {
        g3 := Gui(, editId ? "Edit Shortcut" : "Add Shortcut")
        g3.MarginX := 12, g3.MarginY := 12
        g3.SetFont("s9", "Segoe UI")
        g3.Add("Text",, "Name:")
        eName := g3.Add("Edit", "w420")
        g3.Add("Text",, "Hotkey (e.g., ^!+F1 for Ctrl+Alt+Shift+F1):")
        eHK := g3.Add("Edit", "w200")
        g3.Add("Text",, "Path (file or folder):")
        ePath := g3.Add("Edit", "w420")
        btn := g3.Add("Button",, "Browse")
        btn.OnEvent("Click", (*) => (p := FileSelect("1", , "Select a file", "All Files (*.*)"), p := p ? p : DirSelect("", 1, "Select a folder"), p ? ePath.Value := p : ""))
        btn := g3.Add("Button", "Default", "Save")
        btn.OnEvent("Click", (*) => SaveAndClose())
        btn := g3.Add("Button",, "Cancel")
        btn.OnEvent("Click", (*) => g3.Destroy())

        if editId {
            eName.Value := IniRead(INI, "Shortcuts", editId "_Name",   "")
            eHK.Value   := IniRead(INI, "Shortcuts", editId "_Hotkey", "")
            ePath.Value := IniRead(INI, "Shortcuts", editId "_Path",   "")
        }

        g3.Show("w500 h220")

        SaveAndClose() {
            name := Trim(eName.Value), hk := Trim(eHK.Value), pth := Trim(ePath.Value)
            if (name = "" || hk = "" || pth = "") {
                MsgBox("Please fill in Name, Hotkey, and Path.", "Missing Info", "Icon!")
                return
            }
            idx := editId ? editId : 1
            if !editId
                while (IniRead(INI, "Shortcuts", idx "_Name", "") != "")
                    idx++
            IniWrite(name, INI, "Shortcuts", idx "_Name")
            IniWrite(hk,   INI, "Shortcuts", idx "_Hotkey")
            IniWrite(pth,  INI, "Shortcuts", idx "_Path")
            InitShortcuts()
            RefreshList()
            g3.Destroy()
        }
    }
}
