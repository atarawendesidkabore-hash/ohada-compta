from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_WORKBOOK = ROOT / "_vendor" / "ExcelWebView2" / "ExcelWebView2.xlsm"

DEFAULT_OUTPUT = ROOT / "OHADA-COMPTA-EXACT-SHAREABLE.xlsm"
DEFAULT_DESKTOP_OUTPUT = Path.home() / "OneDrive" / "Desktop" / "OHADA-COMPTA-EXACT-SHAREABLE.xlsm"

LIVE_SITE_URL = "https://atarawendesidkabore-hash.github.io/ohada-compta/?v=20260406a"
HOST_WINDOW_TITLE = "OHADA_COMPTA_EDGE_HOST"
USERFORM_WINDOW_TITLE = "OHADA_COMPTA_WINDOW"

UNUSED_COMPONENTS = [
    "AppConstants",
    "APIFunctions",
    "WV2Globals",
    "MemoryFunctions",
    "clsWebViewEventHandlers",
    "clsWebViewScriptCompleteHandler",
    "clsWebViewContentHandler",
    "pluginLoader",
    "pluginExampleCls",
    "factory",
    "wv2",
    "wv2Environment",
]


HOST_MODULE_CODE = f'''
Option Explicit

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SW_SHOW As Long = 5
Private Const WM_CLOSE As Long = &H10
Private Const HOST_TITLE As String = "{HOST_WINDOW_TITLE}"
Private Const APP_URL As String = "{LIVE_SITE_URL}"

#If VBA7 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private m_shellPrepared As Boolean
Private m_browserProcessId As Long
Private m_browserWindowHwnd As LongPtr

Public Sub LaunchOhadaCompta()
    On Error GoTo HandleError

    ResetOhadaTrace
    TraceOhada "LaunchOhadaCompta:start"
    PrepareExcelShell

    If UserForm1.Visible = False Then
        UserForm1.Show vbModeless
    End If

    UserForm1.Repaint
    ResizeHostedWindow
    TraceOhada "LaunchOhadaCompta:complete"
    Exit Sub

HandleError:
    TraceOhada "LaunchOhadaCompta:error:" & Err.Number & ":" & Err.Description
End Sub

Public Sub PrepareExcelShell()
    On Error Resume Next

    Application.WindowState = xlMaximized
    Application.Caption = "OHADA COMPTA"
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
    End If

    m_shellPrepared = True
End Sub

Public Sub RestoreExcelShell()
    On Error Resume Next

    StopHostedBrowser

    If m_shellPrepared Then
        Application.Caption = vbNullString
        Application.DisplayFormulaBar = True
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"

        If Not ActiveWindow Is Nothing Then
            ActiveWindow.DisplayWorkbookTabs = True
            ActiveWindow.DisplayHeadings = True
            ActiveWindow.DisplayGridlines = True
            ActiveWindow.DisplayHorizontalScrollBar = True
            ActiveWindow.DisplayVerticalScrollBar = True
        End If
    End If

    m_shellPrepared = False
End Sub

Public Function ResolveBrowserExe() As String
    Dim candidate As Variant
    Dim candidates As Variant

    candidates = Array( _
        Environ$("ProgramFiles") & "\\Google\\Chrome\\Application\\chrome.exe", _
        Environ$("ProgramFiles(x86)") & "\\Google\\Chrome\\Application\\chrome.exe", _
        Environ$("LOCALAPPDATA") & "\\Google\\Chrome\\Application\\chrome.exe", _
        Environ$("ProgramFiles(x86)") & "\\Microsoft\\Edge\\Application\\msedge.exe", _
        Environ$("ProgramFiles") & "\\Microsoft\\Edge\\Application\\msedge.exe")

    For Each candidate In candidates
        If Len(candidate) > 0 Then
            If Dir$(CStr(candidate)) <> vbNullString Then
                ResolveBrowserExe = CStr(candidate)
                Exit Function
            End If
        End If
    Next candidate
End Function

Public Function HostRootPath() As String
    Dim rootPath As String
    rootPath = Environ$("LOCALAPPDATA")
    If Len(rootPath) = 0 Then rootPath = Environ$("TEMP")
    If Len(rootPath) = 0 Then rootPath = ThisWorkbook.Path
    HostRootPath = rootPath & "\\OHADA-Compta\\ShareHost"
End Function

Public Function HostHtmlPath() As String
    EnsureFolderTree HostRootPath
    HostHtmlPath = HostRootPath & "\\ohada-edge-host.html"
End Function

Public Function BrowserProfilePath() As String
    EnsureFolderTree HostRootPath
    BrowserProfilePath = HostRootPath & "\\BrowserProfile"
End Function

Public Sub EnsureFolderTree(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String

    If Len(folderPath) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        EnsureFolderTree parentPath
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Public Function BuildHostHtml() As String
    BuildHostHtml = "<!doctype html>" & vbCrLf & _
        "<html lang=""en"">" & vbCrLf & _
        "<head>" & vbCrLf & _
        "  <meta charset=""utf-8"">" & vbCrLf & _
        "  <meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbCrLf & _
        "  <title>" & HOST_TITLE & "</title>" & vbCrLf & _
        "  <style>" & vbCrLf & _
        "    html, body {{ margin: 0; width: 100%; height: 100%; overflow: hidden; background: #061426; }}" & vbCrLf & _
        "    iframe {{ width: 100%; height: 100%; border: 0; display: block; background: #061426; }}" & vbCrLf & _
        "  </style>" & vbCrLf & _
        "</head>" & vbCrLf & _
        "<body>" & vbCrLf & _
        "  <iframe src=""" & APP_URL & """ allow=""clipboard-read; clipboard-write; fullscreen""></iframe>" & vbCrLf & _
        "</body>" & vbCrLf & _
        "</html>"
End Function

Public Sub EnsureHostHtml()
    Dim fileNo As Integer
    fileNo = FreeFile
    Open HostHtmlPath For Output As #fileNo
    Print #fileNo, BuildHostHtml
    Close #fileNo
End Sub

Public Function HostHtmlUri() As String
    EnsureHostHtml
    HostHtmlUri = "file:///" & Replace(HostHtmlPath, "\\", "/")
End Function

Public Sub StartHostedBrowser()
    Dim browserExe As String
    Dim formHwnd As LongPtr
    Dim commandLine As String

    On Error GoTo HandleError
    TraceOhada "StartHostedBrowser:start"

    formHwnd = FindWindow(vbNullString, "{USERFORM_WINDOW_TITLE}")
    If formHwnd = 0 Then Err.Raise 5, "modOhadaShareHost.StartHostedBrowser", "UserForm handle unavailable"

    If m_browserWindowHwnd <> 0 Then
        If IsWindow(m_browserWindowHwnd) <> 0 Then
            AttachHostedWindow formHwnd, m_browserWindowHwnd
            ResizeHostedWindow
            TraceOhada "StartHostedBrowser:reuse"
            Exit Sub
        End If
    End If

    browserExe = ResolveBrowserExe()
    If Len(browserExe) = 0 Then Err.Raise 53, "modOhadaShareHost.StartHostedBrowser", "Chrome or Edge is not installed"

    EnsureHostHtml
    StopHostedBrowser

    commandLine = """" & browserExe & """ --app=""" & HostHtmlUri & """ --user-data-dir=""" & BrowserProfilePath & """ --new-window --disable-session-crashed-bubble --disable-features=msEdgeSidebarV2"
    TraceOhada "StartHostedBrowser:exe=" & browserExe
    m_browserProcessId = Shell(commandLine, vbNormalFocus)
    TraceOhada "StartHostedBrowser:pid=" & m_browserProcessId

    m_browserWindowHwnd = WaitForHostedWindow(15000)
    TraceOhada "StartHostedBrowser:hwnd=" & CStr(m_browserWindowHwnd)
    If m_browserWindowHwnd = 0 Then Err.Raise 5, "modOhadaShareHost.StartHostedBrowser", "Hosted browser window not found"

    AttachHostedWindow formHwnd, m_browserWindowHwnd
    ResizeHostedWindow
    TraceOhada "StartHostedBrowser:complete"
    Exit Sub

HandleError:
    TraceOhada "StartHostedBrowser:error:" & Err.Number & ":" & Err.Description
End Sub

Public Sub StopHostedBrowser()
    On Error Resume Next

    If m_browserWindowHwnd <> 0 Then
        If IsWindow(m_browserWindowHwnd) <> 0 Then
            PostMessage m_browserWindowHwnd, WM_CLOSE, 0, 0
            Sleep 300
        End If
    End If

    Shell "cmd /c taskkill /FI ""WINDOWTITLE eq " & HOST_TITLE & "*"" /T /F", vbHide
    If m_browserProcessId <> 0 Then
        Shell "cmd /c taskkill /PID " & CStr(m_browserProcessId) & " /T /F", vbHide
    End If

    m_browserWindowHwnd = 0
    m_browserProcessId = 0
End Sub

Public Sub ResizeHostedWindow()
    If m_browserWindowHwnd = 0 Then Exit Sub
    If IsWindow(m_browserWindowHwnd) = 0 Then Exit Sub

    MoveWindow m_browserWindowHwnd, 0, 0, CLng(UserForm1.InsideWidth), CLng(UserForm1.InsideHeight), 1
    ShowWindow m_browserWindowHwnd, SW_SHOW
End Sub

Private Sub AttachHostedWindow(ByVal parentHwnd As LongPtr, ByVal childHwnd As LongPtr)
    Dim style As LongPtr

    If parentHwnd = 0 Or childHwnd = 0 Then Exit Sub

    style = GetWindowLongPtr(childHwnd, GWL_STYLE)
    style = style And Not WS_CAPTION
    style = style And Not WS_SYSMENU
    style = style And Not WS_THICKFRAME
    style = style And Not WS_POPUP
    style = style Or WS_CHILD
    style = style Or WS_VISIBLE

    SetParent childHwnd, parentHwnd
    SetWindowLongPtr childHwnd, GWL_STYLE, style
    SetWindowPos childHwnd, 0, 0, 0, CLng(UserForm1.InsideWidth), CLng(UserForm1.InsideHeight), SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_SHOWWINDOW
    ShowWindow childHwnd, SW_SHOW
End Sub

Private Function WaitForHostedWindow(ByVal timeoutMs As Long) As LongPtr
    Dim startedAt As Double
    Dim hwnd As LongPtr

    startedAt = Timer
    Do
        hwnd = FindWindow(vbNullString, HOST_TITLE)
        If hwnd <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                WaitForHostedWindow = hwnd
                Exit Function
            End If
        End If

        If m_browserWindowHwnd <> 0 Then
            If IsWindow(m_browserWindowHwnd) <> 0 Then
                WaitForHostedWindow = m_browserWindowHwnd
                Exit Function
            End If
        End If

        DoEvents
        Sleep 100
        If Timer < startedAt Then startedAt = Timer
    Loop While ((Timer - startedAt) * 1000#) < timeoutMs
End Function

Public Sub MakeWindowFrameless(ByVal formCaption As String)
    Dim hwnd As LongPtr
    Dim style As LongPtr

    On Error Resume Next
    hwnd = FindWindow(vbNullString, formCaption)
    If hwnd = 0 Then Exit Sub

    style = GetWindowLongPtr(hwnd, GWL_STYLE)
    style = style And Not WS_CAPTION
    style = style And Not WS_SYSMENU
    style = style And Not WS_THICKFRAME

    SetWindowLongPtr hwnd, GWL_STYLE, style
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Sub ResetOhadaTrace()
    Dim fileNo As Integer
    fileNo = FreeFile
    Open TracePath For Output As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | trace reset"
    Close #fileNo
End Sub

Public Sub TraceOhada(ByVal message As String)
    Dim fileNo As Integer
    fileNo = FreeFile
    Open TracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fileNo
End Sub

Public Function TracePath() As String
    TracePath = Environ$("TEMP") & "\\OHADA_COMPTA_SHARE_TRACE.log"
End Function
'''.strip() + "\n"


USERFORM1_CODE = f"""
Option Explicit

Private m_appModeReady As Boolean
Private m_browserStarted As Boolean

Private Sub UserForm_Initialize()
    On Error Resume Next
    ConfigureOhadaAppWindow
End Sub

Private Sub UserForm_Activate()
    On Error GoTo HandleError
    ConfigureOhadaAppWindow

    If Not m_browserStarted Then
        DoEvents
        Me.Repaint
        pluginExample.StartHostedBrowser
        m_browserStarted = True
    Else
        pluginExample.ResizeHostedWindow
    End If
    Exit Sub

HandleError:
    pluginExample.TraceOhada "UserForm_Activate:error:" & Err.Number & ":" & Err.Description
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    ConfigureOhadaAppWindow
    pluginExample.ResizeHostedWindow
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    pluginExample.RestoreExcelShell
    Unload Me
End Sub

Private Sub ConfigureOhadaAppWindow()
    On Error Resume Next

    If Not m_appModeReady Then
        Me.Caption = "{USERFORM_WINDOW_TITLE}"
        cmdBack.Visible = False
        cmdForward.Visible = False
        txtUrl.Visible = False
        cmdNewTab.Visible = False
        cmdStopReload.Visible = False
        CommandButton10.Visible = False
        CommandButton7.Visible = False
        browserTabs.Visible = False
        m_appModeReady = True
    End If

    Me.StartUpPosition = 0
    Me.Left = 0
    Me.Top = 0
    Me.Width = Application.UsableWidth
    Me.Height = Application.UsableHeight

    pluginExample.MakeWindowFrameless Me.Caption
    pluginExample.ResizeHostedWindow
End Sub
""".strip() + "\n"


THISWORKBOOK_CODE = """
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo HandleError
    pluginExample.LaunchOhadaCompta
    Exit Sub

HandleError:
    pluginExample.TraceOhada "Workbook_Open:error:" & Err.Number & ":" & Err.Description
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    pluginExample.RestoreExcelShell
End Sub
""".strip() + "\n"


def remove_reference_by_name(vb_project, name: str) -> None:
    to_remove = []
    for reference in vb_project.References:
        try:
            if reference.Name == name:
                to_remove.append(reference)
        except Exception:
            pass

    for reference in to_remove:
        try:
            vb_project.References.Remove(reference)
        except Exception:
            pass


def remove_component_if_exists(vb_project, name: str) -> None:
    try:
        component = vb_project.VBComponents(name)
    except Exception:
        return

    try:
        vb_project.VBComponents.Remove(component)
    except Exception:
        module = component.CodeModule
        if module.CountOfLines:
            module.DeleteLines(1, module.CountOfLines)


def ensure_standard_module(vb_project, name: str):
    try:
        component = vb_project.VBComponents(name)
    except Exception:
        component = vb_project.VBComponents.Add(1)
        component.Name = name
    return component


def replace_component_code(component, code: str) -> None:
    module = component.CodeModule
    if module.CountOfLines:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(code)


def format_start_sheet(workbook) -> None:
    sheet = workbook.Worksheets(1)
    sheet.Name = "START"
    sheet.Cells.Clear()

    sheet.Range("A1").Value = "OHADA COMPTA"
    sheet.Range("A2").Value = "Version Excel partageable du site"
    sheet.Range("A4").Value = "Ce fichier .xlsm peut etre partage seul."
    sheet.Range("A5").Value = "Si Excel affiche un bandeau rouge de securite, faites clic droit sur le fichier > Proprietes > Debloquer, puis rouvrez-le."
    sheet.Range("A6").Value = "Le destinataire doit aussi activer les macros et disposer de Chrome ou Edge."
    sheet.Range("A8").Value = "URL chargee dans Excel :"
    sheet.Range("A9").Value = LIVE_SITE_URL
    sheet.Range("A11").Value = "Cette edition n'a plus besoin des fichiers WebView2_edit.tlb, WebView2Loader.dll ni d'un fichier HTML voisin."

    sheet.Range("A1").Font.Size = 22
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A2").Font.Size = 12
    sheet.Range("A4:A11").Font.Size = 11
    sheet.Range("A9").Font.Color = 0xAA6C00
    sheet.Range("A9").Font.Underline = True
    sheet.Hyperlinks.Add(Anchor=sheet.Range("A9"), Address=LIVE_SITE_URL, TextToDisplay=LIVE_SITE_URL)

    for column in ["A", "B", "C", "D", "E", "F"]:
        sheet.Columns(column).ColumnWidth = 26

    sheet.Range("A1:F20").Interior.Color = 0x0B110A
    sheet.Range("A1:F20").Font.Color = 0xE6E6E6
    sheet.Range("A1:F20").VerticalAlignment = -4160
    sheet.Range("A1:F20").HorizontalAlignment = -4131
    sheet.Rows("1:20").RowHeight = 24
    sheet.Range("A1:F20").WrapText = True


def build_workbook(target_workbook: Path) -> None:
    target_workbook = target_workbook.resolve()
    target_workbook.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE_WORKBOOK, target_workbook)

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 3
    excel.EnableEvents = False

    try:
        workbook = excel.Workbooks.Open(str(target_workbook))
        try:
            excel.VBE.MainWindow.Visible = False
        except Exception:
            pass

        vb_project = workbook.VBProject
        remove_reference_by_name(vb_project, "WebView2_edit")

        for component_name in UNUSED_COMPONENTS:
            remove_component_if_exists(vb_project, component_name)

        replace_component_code(vb_project.VBComponents("ThisWorkbook"), THISWORKBOOK_CODE)
        replace_component_code(vb_project.VBComponents("UserForm1"), USERFORM1_CODE)
        replace_component_code(ensure_standard_module(vb_project, "pluginExample"), HOST_MODULE_CODE)

        format_start_sheet(workbook)
        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        excel.EnableEvents = True
        excel.Quit()
        pythoncom.CoUninitialize()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a single-file shareable Excel workbook that opens the live OHADA-Compta site inside Excel.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--desktop-output", type=Path, default=DEFAULT_DESKTOP_OUTPUT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    build_workbook(args.output)
    if args.desktop_output.resolve() != args.output.resolve():
        build_workbook(args.desktop_output)
    print(f"BUILT:{args.output}")
    print(f"BUILT:{args.desktop_output}")


if __name__ == "__main__":
    main()
