from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[1]
VENDOR_DIR = ROOT / "_vendor" / "ExcelWebView2"
VENDOR_SRC = VENDOR_DIR / "src"
TEMPLATE_WORKBOOK = VENDOR_DIR / "ExcelWebView2.xlsm"
VENDOR_TLB = VENDOR_DIR / "WebView2_edit.tlb"
VENDOR_DLL = VENDOR_DIR / "WebView2Loader.dll"
SYSTEM_DLL_CANDIDATES = [
    Path(r"C:\Program Files\Microsoft Office\root\Office16\WebView2Loader.dll"),
    Path(r"C:\Program Files\Microsoft OneDrive\WebView2Loader.dll"),
]

DEFAULT_OUTPUT = ROOT / "OHADA-COMPTA-EXACT.xlsm"
DEFAULT_DESKTOP_DIR = Path.home() / "OneDrive" / "Desktop" / "OHADA-COMPTA-EXACT"
DEFAULT_DESKTOP_OUTPUT = DEFAULT_DESKTOP_DIR / "OHADA-COMPTA-EXACT.xlsm"

LIVE_SITE_URL = "https://atarawendesidkabore-hash.github.io/ohada-compta/?v=20260406a"
USERDATA_PATH = str((Path.home() / "AppData" / "Local" / "OHADA-Compta" / "WebView2" / "UserData").resolve())
EDGE_HOST_FILE = "ohada-edge-host.html"
EDGE_HOST_HTML = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>OHADA_COMPTA_EDGE_HOST</title>
  <style>
    html, body {{
      margin: 0;
      width: 100%;
      height: 100%;
      overflow: hidden;
      background: #061426;
    }}
    iframe {{
      width: 100%;
      height: 100%;
      border: 0;
      display: block;
      background: #061426;
    }}
  </style>
</head>
<body>
  <iframe src="{LIVE_SITE_URL}" allow="clipboard-read; clipboard-write; fullscreen"></iframe>
</body>
</html>
"""


def resolve_loader_dll() -> Path:
    for candidate in SYSTEM_DLL_CANDIDATES:
        if candidate.exists():
            return candidate
    return VENDOR_DLL


MOD_OHADA_HOST_CODE = '''
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
Private Const EDGE_HOST_WINDOW_TITLE As String = "OHADA_COMPTA_EDGE_HOST"

#If VBA7 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private m_shellPrepared As Boolean
Private m_edgeProcessId As Long
Private m_edgeWindowHwnd As LongPtr
Private m_enumProcessId As Long
Private m_enumWindowHwnd As LongPtr

Public Function OhadaHomePageUrl() As String
    OhadaHomePageUrl = "__LIVE_SITE_URL__"
End Function

Public Function OhadaUserDataPath() As String
    Dim rootPath As String
    rootPath = Environ$("LOCALAPPDATA")
    If Len(rootPath) = 0 Then rootPath = ThisWorkbook.Path
    rootPath = rootPath & "\\OHADA-Compta\\WebView2\\UserData"
    EnsureFolderTree rootPath
    OhadaUserDataPath = rootPath
End Function

Public Function OhadaEdgeHostPath() As String
    OhadaEdgeHostPath = ThisWorkbook.Path & "\\__EDGE_HOST_FILE__"
End Function

Public Function OhadaEdgeHostUri() As String
    OhadaEdgeHostUri = "file:///" & Replace(OhadaEdgeHostPath, "\\", "/")
End Function

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

Public Sub LaunchOhadaCompta()
    On Error GoTo HandleLaunchError

    ResetOhadaTrace
    TraceOhada "LaunchOhadaCompta:start"

    ChDrive Left$(ThisWorkbook.Path, 1)
    ChDir ThisWorkbook.Path

    PrepareExcelShell
    TraceOhada "LaunchOhadaCompta:shell prepared"

    If UserForm1.Visible = False Then
        TraceOhada "LaunchOhadaCompta:show form"
        UserForm1.Show vbModeless
    End If

    UserForm1.Repaint
    ResizeOhadaHost
    TraceOhada "LaunchOhadaCompta:complete"
    Exit Sub

HandleLaunchError:
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

    StopEmbeddedEdge

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

Public Sub MakeOhadaWindowFrameless(ByVal formCaption As String)
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

Public Sub ResizeOhadaHost()
    On Error Resume Next
    ResizeEmbeddedEdgeWindow
End Sub

Public Function BrowserReady(Optional ByVal timeoutSeconds As Double = 20) As Boolean
    Dim startedAt As Double
    startedAt = Timer

    Do
        DoEvents
        If m_edgeWindowHwnd <> 0 And IsWindow(m_edgeWindowHwnd) <> 0 Then
            BrowserReady = True
            Exit Function
        End If

        If Timer < startedAt Then startedAt = Timer
    Loop While (Timer - startedAt) < timeoutSeconds
End Function

Public Function GetBrowserStatus() As String
    If m_edgeWindowHwnd = 0 Then
        GetBrowserStatus = "NO_BROWSER"
        Exit Function
    End If

    GetBrowserStatus = "EDGE_EMBEDDED"
End Function

Public Sub StartEmbeddedEdge()
    Dim edgeExe As String
    Dim formHwnd As LongPtr
    Dim commandLine As String

    On Error GoTo HandleError

    TraceOhada "StartEmbeddedEdge:start"
    formHwnd = FindWindow(vbNullString, UserForm1.Caption)
    If formHwnd = 0 Then Err.Raise 5, "pluginExample.StartEmbeddedEdge", "UserForm handle unavailable"

    If m_edgeWindowHwnd <> 0 Then
        If IsWindow(m_edgeWindowHwnd) <> 0 Then
            TraceOhada "StartEmbeddedEdge:reuse window"
            AttachEdgeWindow formHwnd, m_edgeWindowHwnd
            ResizeEmbeddedEdgeWindow
            Exit Sub
        End If
    End If

    edgeExe = ResolveBrowserExe()
    If Len(edgeExe) = 0 Then Err.Raise 53, "pluginExample.StartEmbeddedEdge", "No supported browser executable found"

    TerminateEdgeHostWindows
    commandLine = """" & edgeExe & """ --app=""" & OhadaEdgeHostUri() & """ --user-data-dir=""" & OhadaUserDataPath() & "\\EdgeApp"" --new-window --disable-session-crashed-bubble --disable-features=msEdgeSidebarV2"
    TraceOhada "StartEmbeddedEdge:exe=" & edgeExe
    m_edgeProcessId = Shell(commandLine, vbNormalFocus)
    TraceOhada "StartEmbeddedEdge:pid=" & m_edgeProcessId

    m_edgeWindowHwnd = WaitForEdgeWindow(15000)
    TraceOhada "StartEmbeddedEdge:hwnd=" & m_edgeWindowHwnd
    If m_edgeWindowHwnd = 0 Then Err.Raise 5, "pluginExample.StartEmbeddedEdge", "Edge window not found"

    AttachEdgeWindow formHwnd, m_edgeWindowHwnd
    ResizeEmbeddedEdgeWindow
    TraceOhada "StartEmbeddedEdge:complete"
    Exit Sub

HandleError:
    TraceOhada "StartEmbeddedEdge:error:" & Err.Number & ":" & Err.Description
End Sub

Public Sub StopEmbeddedEdge()
    On Error Resume Next

    If m_edgeWindowHwnd <> 0 Then
        If IsWindow(m_edgeWindowHwnd) <> 0 Then
            PostMessage m_edgeWindowHwnd, WM_CLOSE, 0, 0
            Sleep 300
        End If
    End If

    TerminateEdgeHostWindows

    If m_edgeProcessId <> 0 Then
        Shell "cmd /c taskkill /PID " & CStr(m_edgeProcessId) & " /T /F", vbHide
    End If

    m_edgeWindowHwnd = 0
    m_edgeProcessId = 0
End Sub

Public Sub ResizeEmbeddedEdgeWindow()
    If m_edgeWindowHwnd = 0 Then Exit Sub
    If IsWindow(m_edgeWindowHwnd) = 0 Then Exit Sub

    MoveWindow m_edgeWindowHwnd, 0, 0, CLng(UserForm1.InsideWidth), CLng(UserForm1.InsideHeight), 1
    ShowWindow m_edgeWindowHwnd, SW_SHOW
End Sub

Private Sub AttachEdgeWindow(ByVal parentHwnd As LongPtr, ByVal childHwnd As LongPtr)
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

Private Sub TerminateEdgeHostWindows()
    On Error Resume Next
    Shell "cmd /c taskkill /FI ""WINDOWTITLE eq " & EDGE_HOST_WINDOW_TITLE & "*"" /T /F", vbHide
    Sleep 300
End Sub

Private Function WaitForEdgeWindow(ByVal timeoutMs As Long) As LongPtr
    Dim startedAt As Double
    Dim hwnd As LongPtr

    startedAt = Timer

    Do
        hwnd = FindWindow(vbNullString, EDGE_HOST_WINDOW_TITLE)
        If hwnd <> 0 Then
            If IsWindowVisible(hwnd) <> 0 Then
                WaitForEdgeWindow = hwnd
                Exit Function
            End If
        End If

        If m_edgeWindowHwnd <> 0 Then
            If IsWindow(m_edgeWindowHwnd) <> 0 Then
                WaitForEdgeWindow = m_edgeWindowHwnd
                Exit Function
            End If
        End If

        If hwnd <> 0 Then
            WaitForEdgeWindow = hwnd
            Exit Function
        End If

        DoEvents
        Sleep 100
        If Timer < startedAt Then startedAt = Timer
    Loop While ((Timer - startedAt) * 1000#) < timeoutMs
End Function

Public Sub ResetOhadaTrace()
    Dim tracePath As String
    Dim fileNo As Integer

    tracePath = OhadaTracePath()
    fileNo = FreeFile
    Open tracePath For Output As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | trace reset"
    Close #fileNo
End Sub

Public Sub TraceOhada(ByVal message As String)
    Dim tracePath As String
    Dim fileNo As Integer

    tracePath = OhadaTracePath()
    fileNo = FreeFile
    Open tracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fileNo
End Sub

Public Function OhadaTracePath() As String
    OhadaTracePath = Environ$("TEMP") & "\\OHADA_COMPTA_TRACE.log"
End Function
'''.strip().replace("__LIVE_SITE_URL__", LIVE_SITE_URL).replace("__EDGE_HOST_FILE__", EDGE_HOST_FILE) + "\n"


USERFORM1_CODE = """
Option Explicit

Private m_appModeReady As Boolean
Private m_browserStarted As Boolean

Private Sub browserTabs_Change()
    If (Not Not g_wv2) <> 0 Then
        g_wv2(browserTabs.SelectedItem.Index).Focus
        g_selectedTabIndex = browserTabs.SelectedItem.Index
    End If
End Sub

Private Sub cmdBack_Click()
    ActiveBrowserTab.GoBack
End Sub

Private Sub cmdForward_Click()
    ActiveBrowserTab.GoForward
End Sub

Private Sub cmdNewTab_Click()
    pluginExample.StartEmbeddedEdge
End Sub

Private Sub cmdStopReload_Click()
    Dim i As Long
    If cmdStopReload.Caption = "X" Then
        For i = 0 To UBound(g_wv2)
            g_wv2(i).StopLoading
        Next i
    Else
        ActiveBrowserTab.Reload
    End If
End Sub

Private Sub CommandButton7_Click()
    g_wv2(browserTabs.SelectedItem.Index).OpenDevTools
End Sub

Private Sub CommandButton10_Click()
    g_wv2(browserTabs.SelectedItem.Index).OpenDevTools
End Sub

Private Sub txtUrl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ActiveBrowserTab.OpenUrl txtUrl.Text
        txtUrl.SelStart = 0
        txtUrl.SelLength = Len(txtUrl.Text)
    End If
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo HandleError
    pluginExample.TraceOhada "UserForm_Initialize:start"
    ConfigureOhadaAppWindow
    pluginExample.TraceOhada "UserForm_Initialize:complete"
    Exit Sub

HandleError:
    pluginExample.TraceOhada "UserForm_Initialize:error:" & Err.Number & ":" & Err.Description
End Sub

Private Sub UserForm_Activate()
    On Error GoTo HandleError
    pluginExample.TraceOhada "UserForm_Activate:start"
    ConfigureOhadaAppWindow
    If Not m_browserStarted Then
        DoEvents
        UserForm1.Repaint
        pluginExample.StartEmbeddedEdge
        m_browserStarted = True
    End If
    pluginExample.TraceOhada "UserForm_Activate:complete"
    Exit Sub

HandleError:
    pluginExample.TraceOhada "UserForm_Activate:error:" & Err.Number & ":" & Err.Description
End Sub

Private Sub UserForm_Resize()
    On Error GoTo HandleError
    pluginExample.ResizeOhadaHost
    Exit Sub

HandleError:
    pluginExample.TraceOhada "UserForm_Resize:error:" & Err.Number & ":" & Err.Description
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    WV2Globals.CleanUp
    pluginExample.StopEmbeddedEdge
    pluginExample.RestoreExcelShell
    Unload Me
End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)

End Sub

Private Sub ConfigureOhadaAppWindow()
    On Error GoTo HandleError
    pluginExample.TraceOhada "ConfigureOhadaAppWindow:start"

    If Not m_appModeReady Then
        Me.Caption = "OHADA_COMPTA_WINDOW"
        cmdBack.Visible = False
        cmdForward.Visible = False
        txtUrl.Visible = False
        cmdNewTab.Visible = False
        cmdStopReload.Visible = False
        CommandButton10.Visible = False
        CommandButton7.Visible = False
        browserTabs.Visible = False
        browserTabs.Top = -24
        browserTabs.Left = 0
        browserTabs.Width = Me.InsideWidth
        browserTabs.Height = Me.InsideHeight + 24
        m_appModeReady = True
    End If

    Me.StartUpPosition = 0
    Me.Left = 0
    Me.Top = 0
    Me.Width = Application.UsableWidth
    Me.Height = Application.UsableHeight

    pluginExample.MakeOhadaWindowFrameless Me.Caption
    pluginExample.ResizeOhadaHost
    pluginExample.TraceOhada "ConfigureOhadaAppWindow:complete"
    Exit Sub

HandleError:
    pluginExample.TraceOhada "ConfigureOhadaAppWindow:error:" & Err.Number & ":" & Err.Description
End Sub
""".strip() + "\n"


THISWORKBOOK_CODE = """
Option Explicit

Private Sub Workbook_Open()
    ChDrive Left$(ThisWorkbook.Path, 1)
    ChDir ThisWorkbook.Path
    pluginExample.LaunchOhadaCompta
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    pluginExample.RestoreExcelShell
End Sub
""".strip() + "\n"


PLUGIN_LOADER_CODE = """
Option Explicit
'OHADA host build: plugin loading is disabled because this workbook is used as an app shell.

Private m_plugins As pluginManagerSingleton

Public Property Get PluginManager() As pluginManagerSingleton
    If m_plugins Is Nothing Then
        Set m_plugins = New pluginManagerSingleton
    End If
    Set PluginManager = m_plugins
End Property

Public Sub LoadPlugins()
End Sub
""".strip() + "\n"


PLUGIN_EXAMPLE_CLS_CODE = """
Option Explicit

Implements pluginInterface

Private WithEvents m_WebView2Event As clsWebViewEventHandlers
Private WithEvents m_ContentEvent As clsWebViewContentHandler
Private WithEvents m_ScriptEvent As clsWebViewScriptCompleteHandler

Private Sub m_ContentEvent_WebResourceResponseViewGetContentCompleted(res As clsWebResData, Content As WebView2_edit.IStream)
End Sub

Private Sub m_ScriptEvent_wv2ScriptComplete(ByVal sender As wv2, resultObjectAsJson As String)
End Sub

Private Sub m_WebView2Event_DocumentTitleChanged(sender As WebView2_edit.ICoreWebView2, args As Long)
End Sub

Private Sub m_WebView2Event_DOMContentLoaded(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2DOMContentLoadedEventArgs)
End Sub

Private Sub m_WebView2Event_NavigationCompleted(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationCompletedEventArgs)
End Sub

Private Sub m_WebView2Event_NavigationStarting(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2NavigationStartingEventArgs)
End Sub

Private Sub m_WebView2Event_WebResourceRequested(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceRequestedEventArgs)
End Sub

Private Sub m_WebView2Event_WebResourceResponseReceived(sender As WebView2_edit.ICoreWebView2, args As WebView2_edit.ICoreWebView2WebResourceResponseReceivedEventArgs)
End Sub

Private Sub m_WebView2Event_wv2ControllerReady(createdController As WebView2_edit.ICoreWebView2Controller)
End Sub

Private Sub m_WebView2Event_wv2EnvironmentReady(createdEnvironment As WebView2_edit.ICoreWebView2Environment)
End Sub

Private Property Get pluginInterface_NewInstance() As pluginInterface
    Set pluginInterface_NewInstance = New pluginExampleCls
End Property

Private Property Get pluginInterface_ContentEvent() As clsWebViewContentHandler
    Set pluginInterface_ContentEvent = m_ContentEvent
End Property

Private Property Set pluginInterface_ContentEvent(ByVal RHS As clsWebViewContentHandler)
    Set m_ContentEvent = RHS
End Property

Private Property Set pluginInterface_ScriptEvent(ByVal RHS As clsWebViewScriptCompleteHandler)
    Set m_ScriptEvent = RHS
End Property

Private Property Get pluginInterface_ScriptEvent() As clsWebViewScriptCompleteHandler
    Set pluginInterface_ScriptEvent = m_ScriptEvent
End Property

Private Property Set pluginInterface_WebView2Event(ByVal RHS As clsWebViewEventHandlers)
    Set m_WebView2Event = RHS
End Property

Private Property Get pluginInterface_WebView2Event() As clsWebViewEventHandlers
    Set pluginInterface_WebView2Event = m_WebView2Event
End Property
""".strip() + "\n"


FACTORY_CODE = """
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'factory
'***

Public Function NewTab() As wv2
    On Error GoTo HandleError
    pluginExample.TraceOhada "factory.NewTab:start"
    Set NewTab = New wv2
    pluginExample.TraceOhada "factory.NewTab:complete"
    Exit Function

HandleError:
    pluginExample.TraceOhada "factory.NewTab:error:" & Err.Number & ":" & Err.Description
End Function
""".strip() + "\n"


APIFUNCTIONS_CODE = """
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'APIFunctions
'64-bit compatible declarations for Office VBA7
'***

#If VBA7 Then
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserExecutableFolder As LongPtr, ByVal userDataFolder As LongPtr, ByVal environmentOptions As LongPtr, ByVal createdEnvironmentCallback As LongPtr) As Long

Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function SetMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal hMenu As LongPtr) As Long
Public Declare PtrSafe Function CreateMenu Lib "user32" () As LongPtr
Public Declare PtrSafe Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As LongPtr) As Long
Public Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Public Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Public Declare PtrSafe Function CoInitialize Lib "ole32" (ByRef pvReserved As Any) As Long
Public Declare PtrSafe Function CoInitializeEx Lib "ole32" (ByVal pvReserved As LongPtr, ByVal dwCoInit As Long) As Long
Public Declare PtrSafe Sub CoUninitialize Lib "ole32" ()
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As LongPtr, hGlobal As Any) As Long

Public Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long

Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Public Declare PtrSafe Function GetAddrOf Lib "kernel32" Alias "MulDiv" (nNumber As Any, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As LongPtr
Public Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As LongPtr)
Public Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare PtrSafe Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare PtrSafe Function timeGetTime Lib "winmm" () As Long

Public Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As LongPtr, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Public Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As LongPtr) As Long

Public Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As LongPtr) As Long
#Else
Public Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal browserExecutableFolder As Long, ByVal userDataFolder As Long, ByVal environmentOptions As Long, ByVal createdEnvironmentCallback As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function CoInitialize Lib "ole32" (ByRef pvReserved As Any) As Long
Public Declare Function CoInitializeEx Lib "ole32" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long
Public Declare Sub CoUninitialize Lib "ole32" ()
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Public Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Any) As Long
Public Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgvt As Any, prgpvarg As Any, pvargResult As Variant) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function GetAddrOf Lib "kernel32" Alias "MulDiv" (nNumber As Any, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, Image As Long) As Long
Public Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
#End If
""".strip() + "\n"


WV2_ENVIRONMENT_CODE = """
Option Explicit
'ExcelWebView2 by Lucas Plumb @ 2023
'WebView2Environment object, do NOT instantiate this manually!
'simply create a new wv2 class object using:
'    Dim wv As wv2
'    Set wv = New wv2
'the wv2 object itself will handle everything needed for setup
'this class object will automatically be instantiated if needed by the wv2 object

Private m_WebView2Environment As ICoreWebView2Environment
Private WithEvents m_webViewHandlers As clsWebViewEventHandlers
Attribute m_webViewHandlers.VB_VarHelpID = -1

Public Event wv2CtrlReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)
Public Event wv2EnvReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)
Public Event wv2Ready(ByRef env As wv2Environment)


Public Property Get this() As ICoreWebView2Environment
    Dim userDataDir As String

    If m_WebView2Environment Is Nothing Then
        TraceOhada "wv2Environment.this:create env"
        userDataDir = OhadaUserDataPath()
        If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userDataDir), 0&, handler) <> S_OK Then
            Unload Me
        Else
        End If
    Else
        Set this = m_WebView2Environment
    End If
End Property

Public Sub Init()
    Dim userDataDir As String

    TraceOhada "wv2Environment.Init:start"
    userDataDir = OhadaUserDataPath()
    If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userDataDir), 0&, handler) <> S_OK Then
        Unload Me
    Else
    End If
    RaiseEvent wv2Ready(Me)
    TraceOhada "wv2Environment.Init:complete"
End Sub

Public Property Get handler() As clsWebViewEventHandlers
    If m_webViewHandlers Is Nothing Then
        Set m_webViewHandlers = New clsWebViewEventHandlers
    End If
    Set handler = m_webViewHandlers
End Property

Public Property Set handler(m_handler As clsWebViewEventHandlers)
    Set m_webViewHandlers = m_handler
End Property

Private Sub Class_Initialize()
    If Not g_wv2Env Is Nothing Then
        MsgBox "wv2Environment error - class created when environment already exists", vbCritical, "Error"
    End If
    Set g_wv2Env = Me
End Sub

Private Sub Class_Terminate()
    Set m_WebView2Environment = Nothing
    Set m_webViewHandlers = Nothing
End Sub

Private Sub m_webViewHandlers_wv2ControllerReady(ByRef createdController As WebView2_edit.ICoreWebView2Controller)
    Debug.Print "controller ready in env"
    TraceOhada "wv2Environment.controller ready"
    RaiseEvent wv2CtrlReady(createdController)
End Sub

Private Sub m_webViewHandlers_wv2EnvironmentReady(ByRef createdEnvironment As WebView2_edit.ICoreWebView2Environment)
    Set m_WebView2Environment = createdEnvironment
    Set g_Env = createdEnvironment
    RaiseEvent wv2EnvReady(g_Env)
    Debug.Print "environment ready in env"
    TraceOhada "wv2Environment.environment ready"
End Sub
""".strip() + "\n"


def vendor_text(name: str) -> str:
    return (VENDOR_SRC / name).read_text(encoding="utf-8-sig")


def vba_code_only(source: str) -> str:
    lines = source.splitlines()
    for index, line in enumerate(lines):
        if line.strip() == "Option Explicit":
            cleaned = []
            for kept in lines[index:]:
                if kept.lstrip().startswith("Attribute "):
                    continue
                cleaned.append(kept)
            return "\n".join(cleaned) + "\n"
    return source


def patched_app_constants_code() -> str:
    code = vba_code_only(vendor_text("AppConstants.bas"))
    code = code.replace(
        'Public Const userdata = "C:\\ExcelWebView2\\userdata\\" \'set a profile/userdata folder to the browser to use',
        f'Public Const userdata = "{USERDATA_PATH}\\\\" \'set a profile/userdata folder to the browser to use',
    )
    code = code.replace(
        'Public Const homePageUrl = "https://google.com"',
        f'Public Const homePageUrl = "{LIVE_SITE_URL}"',
    )
    return code


def patched_wv2_globals_code() -> str:
    code = vba_code_only(vendor_text("WV2Globals.bas"))
    code = code.replace("Public g_webHostHwnd As Long", "Public g_webHostHwnd As LongPtr")
    return code


def patched_memory_functions_code() -> str:
    return """
Option Explicit
'***
'ExcelWebView2 by Lucas Plumb @ 2023
'MemoryFunctions
'64-bit adjusted pointer helpers
'***

Public Function GetBaseAddress(vb_array As Variant) As LongPtr
    Dim vType As Integer
    Dim lp As LongPtr
    Dim address As LongPtr

    CopyMemory vType, vb_array, 2
    CopyMemory lp, ByVal VarPtr(vb_array) + 8, LenB(lp)

    If (vType And VT_BY_REF) <> 0 Then
        CopyMemory lp, ByVal lp, LenB(lp)
        CopyMemory address, ByVal lp + 16, LenB(address)
        GetBaseAddress = address
    End If
End Function

Public Function StrFromPtr(ByVal lpStr As LongPtr) As String
    Dim bStr() As Byte
    Dim cChars As Long

    On Error Resume Next
    cChars = lstrlen(lpStr) * 2
    If cChars > 0 Then
        ReDim bStr(0 To cChars - 1) As Byte
        CopyMemory bStr(0), ByVal lpStr, cChars
    End If
    StrFromPtr = bStr
End Function

Private Function GetStrFromPtrW(ByVal Ptr As LongPtr) As String
    SysReAllocString VarPtr(GetStrFromPtrW), Ptr
End Function

Public Function shr(ByVal value As Long, ByVal Shift As Byte) As Long
    shr = value
    If Shift > 0 Then
        shr = Int(shr / (2 ^ Shift))
    End If
End Function

Public Function shl(ByVal value As Long, ByVal Shift As Byte) As Long
    shl = value
    If Shift > 0 Then
        Dim i As Byte
        Dim m As Long
        For i = 1 To Shift
            m = shl And &H40000000
            shl = (shl And &H3FFFFFFF) * 2
            If m <> 0 Then
                shl = shl Or &H80000000
            End If
        Next i
    End If
End Function

Function hresToWin32(e As Long) As Long
    If e <= 0 Then hresToWin32 = e: Exit Function
    Dim s1 As Long, s2 As Long, s3 As Long
    s1 = e And &HFFFF&
    s2 = shl(&H7&, 16)
    s3 = &H80000000
    hresToWin32 = s1 Or s2 Or s3
End Function
""".strip() + "\n"


def patched_cls_webview_event_handlers_code() -> str:
    code = vba_code_only(vendor_text("clsWebViewEventHandlers.cls"))
    return code + """

Public Property Get AsControllerCompletedHandler() As ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
    Set AsControllerCompletedHandler = Me
End Property

Public Property Get AsEnvironmentCompletedHandler() As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    Set AsEnvironmentCompletedHandler = Me
End Property

Public Property Get EnvironmentCompletedHandlerPtr() As LongPtr
    Dim callbackHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    Set callbackHandler = Me
    EnvironmentCompletedHandlerPtr = ObjPtr(callbackHandler)
End Property

Public Property Get AsWebResourceRequestedHandler() As ICoreWebView2WebResourceRequestedEventHandler
    Set AsWebResourceRequestedHandler = Me
End Property

Public Property Get AsWebResourceResponseReceivedHandler() As ICoreWebView2WebResourceResponseReceivedEventHandler
    Set AsWebResourceResponseReceivedHandler = Me
End Property

Public Property Get AsDOMContentLoadedHandler() As ICoreWebView2DOMContentLoadedEventHandler
    Set AsDOMContentLoadedHandler = Me
End Property

Public Property Get AsNavigationCompletedHandler() As ICoreWebView2NavigationCompletedEventHandler
    Set AsNavigationCompletedHandler = Me
End Property

Public Property Get AsNavigationStartingHandler() As ICoreWebView2NavigationStartingEventHandler
    Set AsNavigationStartingHandler = Me
End Property

Public Property Get AsDocumentTitleChangedHandler() As ICoreWebView2DocumentTitleChangedEventHandler
    Set AsDocumentTitleChangedHandler = Me
End Property
""".strip() + "\n"


def patched_cls_webview_script_complete_handler_code() -> str:
    code = vba_code_only(vendor_text("clsWebViewScriptCompleteHandler.cls"))
    return code + """

Public Property Get AsExecuteScriptCompletedHandler() As ICoreWebView2ExecuteScriptCompletedHandler
    Set AsExecuteScriptCompletedHandler = Me
End Property
""".strip() + "\n"


def patched_cls_webview_content_handler_code() -> str:
    code = vba_code_only(vendor_text("clsWebViewContentHandler.cls"))
    return code + """

Public Property Get AsGetContentCompletedHandler() As ICoreWebView2WebResourceResponseViewGetContentCompletedHandler
    Set AsGetContentCompletedHandler = Me
End Property
""".strip() + "\n"


def patched_wv2_code() -> str:
    code = vba_code_only(vendor_text("wv2.cls"))

    old_frame_block = """            With g_webFrame
                .Top = UserForm1.browserTabs.Top + 14
                .Left = UserForm1.browserTabs.Left
                .Height = UserForm1.browserTabs.Height - 14
                .Width = UserForm1.browserTabs.Width
                .TabStop = False
                .Visible = True
            End With"""
    new_frame_block = """            With g_webFrame
                .Top = 0
                .Left = 0
                .Height = UserForm1.InsideHeight
                .Width = UserForm1.InsideWidth
                .BorderStyle = 0
                .TabStop = False
                .Visible = True
            End With"""
    code = code.replace(old_frame_block, new_frame_block)

    title_anchor = """Public Property Let pageSource(val As String)
    m_pageSource = val
End Property
Public Property Get NavigationComplete() As Boolean"""
    title_insert = """Public Property Let pageSource(val As String)
    m_pageSource = val
End Property
Public Property Get DocumentTitle() As String
    If Not m_WebViewCore Is Nothing Then
        DocumentTitle = StrFromPtr(m_WebViewCore.DocumentTitle)
    End If
End Property
Public Property Get NavigationComplete() As Boolean"""
    code = code.replace(title_anchor, title_insert)

    execute_script_old = """Function ExecuteScript(javaScript As String, Optional PropLet As String = vbNullString)
    Dim scriptHandler As clsWebViewScriptCompleteHandler 'create an instance of the scriptCompleteHandler class and set its parent, which will then call wv2.ExecuteScriptCompletedHandler back to us in this wv2 instance
    Set scriptHandler = New clsWebViewScriptCompleteHandler
    Set scriptHandler.Parent = Me
    scriptHandler.PropLet = PropLet 'if we want the result of the script to set some variable when it completes, we can use this argument
    m_WebViewCore.ExecuteScript javaScript, scriptHandler 'm_scriptHandler
End Function"""
    execute_script_new = """Function ExecuteScript(javaScript As String, Optional PropLet As String = vbNullString)
    Dim scriptHandler As clsWebViewScriptCompleteHandler 'create an instance of the scriptCompleteHandler class and set its parent, which will then call wv2.ExecuteScriptCompletedHandler back to us in this wv2 instance
    Set scriptHandler = New clsWebViewScriptCompleteHandler
    Set scriptHandler.Parent = Me
    scriptHandler.PropLet = PropLet 'if we want the result of the script to set some variable when it completes, we can use this argument
    m_WebViewCore.ExecuteScript javaScript, scriptHandler.AsExecuteScriptCompletedHandler 'm_scriptHandler
End Function"""
    code = code.replace(execute_script_old, execute_script_new)

    webview_ready_old = """Private Sub WebViewReady() 'called each time a new wv2controller is ready (every new tab)
    Dim token As EventRegistrationToken 'just pass the same token pointer around, at this point we dont really care to ever remove these handlers, maybe <TODO> in the future
    
    'initialize web event handlers
    m_WebViewCore.AddWebResourceRequestedFilter "*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL 'you MUST add a WebResource filter to receive the events at all, we want to see all events so use the * wildcard
    m_WebViewCore.add_WebResourceRequested m_webViewHandlers, token
    m_WebViewCore.add_WebResourceResponseReceived m_webViewHandlers, token
    m_WebViewCore.add_DOMContentLoaded m_webViewHandlers, token
    m_WebViewCore.add_NavigationCompleted m_webViewHandlers, token
    m_WebViewCore.add_NavigationStarting m_webViewHandlers, token
    m_WebViewCore.add_DocumentTitleChanged m_webViewHandlers, token
    
    'navigate to homepage
    Me.OpenUrl homePageUrl
End Sub"""
    webview_ready_new = """Private Sub WebViewReady() 'called each time a new wv2controller is ready (every new tab)
    Dim token As EventRegistrationToken 'just pass the same token pointer around, at this point we dont really care to ever remove these handlers, maybe <TODO> in the future
    
    'initialize web event handlers
    m_WebViewCore.AddWebResourceRequestedFilter "*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL 'you MUST add a WebResource filter to receive the events at all, we want to see all events so use the * wildcard
    m_WebViewCore.add_WebResourceRequested m_webViewHandlers.AsWebResourceRequestedHandler, token
    m_WebViewCore.add_WebResourceResponseReceived m_webViewHandlers.AsWebResourceResponseReceivedHandler, token
    m_WebViewCore.add_DOMContentLoaded m_webViewHandlers.AsDOMContentLoadedHandler, token
    m_WebViewCore.add_NavigationCompleted m_webViewHandlers.AsNavigationCompletedHandler, token
    m_WebViewCore.add_NavigationStarting m_webViewHandlers.AsNavigationStartingHandler, token
    m_WebViewCore.add_DocumentTitleChanged m_webViewHandlers.AsDocumentTitleChangedHandler, token
    
    'navigate to homepage
    Me.OpenUrl homePageUrl
End Sub"""
    code = code.replace(webview_ready_old, webview_ready_new)

    get_env_old = """Private Function GetWebView2Env(ByRef m_tab As MSForms.Tab) As wv2Environment
    Set m_tab = UserForm1.browserTabs.Tabs.Add("tab" & myIndex, "New Tab", myIndex)
    
    If g_wv2Env Is Nothing Then
        If g_webFrame Is Nothing Then
            Set g_webFrame = UserForm1.Controls.Add("Forms.Frame.1", "tab_Frame" & myIndex, False)
            With g_webFrame
                .Top = 0
                .Left = 0
                .Height = UserForm1.InsideHeight
                .Width = UserForm1.InsideWidth
                .BorderStyle = 0
                .TabStop = False
                .Visible = True
            End With
        End If
    
        g_webHostHwnd = g_webFrame.[_GethWnd]
    
        Set m_wv2env = New wv2Environment
        m_wv2env.Init
    End If
    
    Set m_wv2env = g_wv2Env
    Set GetWebView2Env = m_wv2env
    'Debug.Print myIndex
    'Set UserForm1.browserTabs.SelectedItem = UserForm1.browserTabs.Tabs.Item(myIndex + 1)
End Function"""
    get_env_new = """Private Function GetWebView2Env(ByRef m_tab As MSForms.Tab) As wv2Environment
    pluginExample.TraceOhada "GetWebView2Env:start"
    Set m_tab = UserForm1.browserTabs.Tabs.Add("tab" & myIndex, "New Tab", myIndex)
    pluginExample.TraceOhada "GetWebView2Env:tab ready"
    
    If g_wv2Env Is Nothing Then
        DoEvents
        UserForm1.Repaint
        g_webHostHwnd = FindWindow(vbNullString, UserForm1.Caption)
        pluginExample.TraceOhada "GetWebView2Env:form hwnd=" & g_webHostHwnd
        If g_webHostHwnd = 0 Then
            Err.Raise 5, "wv2.GetWebView2Env", "UserForm handle unavailable"
        End If
        
        Set m_wv2env = New wv2Environment
        pluginExample.TraceOhada "GetWebView2Env:env created"
        m_wv2env.Init
        pluginExample.TraceOhada "GetWebView2Env:init called"
    End If
    
    Set m_wv2env = g_wv2Env
    Set GetWebView2Env = m_wv2env
    pluginExample.TraceOhada "GetWebView2Env:complete"
End Function"""
    code = code.replace(get_env_old, get_env_new)

    class_init_old = """Private Sub Class_Initialize()
    Dim newCount As Integer
    Set resDict = New Dictionary
    
    'keep a reference to this instance in global
    If (Not Not g_wv2) = 0 Then 'if the g_wv2 array is uninitialized, this is the first instance
        
        'cleanup/initialize plugins on first instance creation
        PluginManager.Kill
        pluginLoader.LoadPlugins
        
        newCount = 0
    Else
        newCount = UBound(g_wv2) + 1
    End If
    myIndex = newCount
    
    If g_wv2Env Is Nothing Then
        'create webview2 environment if it doesnt exist yet
        Set m_wv2env = GetWebView2Env(m_tab)
        'controller will be created automatically when environment is initialized
        'see m_wv2env_wv2EnvReady
    Else
        Set m_wv2env = GetWebView2Env(m_tab)
        'environment already exists, so we need to create a new controller instead (this is called when this class is initialized more than once, for instance creating a new browser tab)
        g_wv2Env.this.CreateCoreWebView2Controller g_webHostHwnd, handler
    End If
    

    
    ReDim Preserve g_wv2(newCount)
    Set g_wv2(newCount) = Me
End Sub"""
    class_init_new = """Private Sub Class_Initialize()
    Dim newCount As Integer
    pluginExample.TraceOhada "wv2.Class_Initialize:start"
    Set resDict = New Dictionary
    pluginExample.TraceOhada "wv2.Class_Initialize:dict ready"
    
    If (Not Not g_wv2) = 0 Then
        pluginExample.TraceOhada "wv2.Class_Initialize:first instance"
        PluginManager.Kill
        pluginLoader.LoadPlugins
        newCount = 0
    Else
        pluginExample.TraceOhada "wv2.Class_Initialize:additional instance"
        newCount = UBound(g_wv2) + 1
    End If
    myIndex = newCount
    pluginExample.TraceOhada "wv2.Class_Initialize:index set"
    
    If g_wv2Env Is Nothing Then
        pluginExample.TraceOhada "wv2.Class_Initialize:request env"
        Set m_wv2env = GetWebView2Env(m_tab)
        pluginExample.TraceOhada "wv2.Class_Initialize:env received"
    Else
        pluginExample.TraceOhada "wv2.Class_Initialize:env exists"
        Set m_wv2env = GetWebView2Env(m_tab)
        pluginExample.TraceOhada "wv2.Class_Initialize:creating controller"
        g_wv2Env.this.CreateCoreWebView2Controller g_webHostHwnd, handler
        pluginExample.TraceOhada "wv2.Class_Initialize:controller created"
    End If
    
    ReDim Preserve g_wv2(newCount)
    Set g_wv2(newCount) = Me
    pluginExample.TraceOhada "wv2.Class_Initialize:complete"
End Sub"""
    code = code.replace(class_init_old, class_init_new)

    code = code.replace(
        "        g_wv2Env.this.CreateCoreWebView2Controller g_webHostHwnd, handler",
        """        Dim hostWindow As Long
        hostWindow = CLng(g_webHostHwnd)
        g_wv2Env.this.CreateCoreWebView2Controller hostWindow, handler.AsControllerCompletedHandler""",
    )
    code = code.replace(
        "        args.Response.GetContent req.contentHandler",
        """        args.Response.GetContent req.contentHandler.AsGetContentCompletedHandler""",
    )

    return code


def patched_wv2_environment_code() -> str:
    code = vba_code_only(vendor_text("wv2Environment.cls"))
    code = code.replace(
        "        If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, handler) <> S_OK Then",
        """        Dim environmentHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
        Set environmentHandler = handler
        If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, ObjPtr(environmentHandler)) <> S_OK Then""",
    )
    code = code.replace(
        "    If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, handler) <> S_OK Then",
        """    Dim environmentHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    Set environmentHandler = handler
    If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, ObjPtr(environmentHandler)) <> S_OK Then""",
    )
    init_old = """Public Sub Init()
    Dim environmentHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    Set environmentHandler = handler
    If CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, ObjPtr(environmentHandler)) <> S_OK Then
        'MessageBox 0, "Failed to create environment", "Error", 0
        Unload Me
    Else
        'we could raise an "environment creation successful" event here, but we do it in "handler" instead
    End If
    RaiseEvent wv2Ready(Me)
End Sub"""
    init_new = """Public Sub Init()
    On Error GoTo HandleError

    Dim callbackPtr As LongPtr
    Dim createResult As Long

    pluginExample.TraceOhada "wv2Environment.Init:start"
    pluginExample.TraceOhada "wv2Environment.Init:handler ready"
    callbackPtr = handler.EnvironmentCompletedHandlerPtr
    pluginExample.TraceOhada "wv2Environment.Init:callback ptr=" & callbackPtr
    createResult = CreateCoreWebView2EnvironmentWithOptions(0&, StrPtr(userdata), 0&, callbackPtr)
    pluginExample.TraceOhada "wv2Environment.Init:result=" & createResult
    If createResult <> S_OK Then
        pluginExample.TraceOhada "wv2Environment.Init:non_s_ok"
        Unload Me
    Else
        pluginExample.TraceOhada "wv2Environment.Init:api ok"
    End If
    RaiseEvent wv2Ready(Me)
    pluginExample.TraceOhada "wv2Environment.Init:complete"
    Exit Sub

HandleError:
    pluginExample.TraceOhada "wv2Environment.Init:error:" & Err.Number & ":" & Err.Description
End Sub"""
    code = code.replace(init_old, init_new)
    code = code.replace(
        '    RaiseEvent wv2EnvReady(g_Env)',
        '    pluginExample.TraceOhada "wv2Environment.Event:env ready"' + "\n" + '    RaiseEvent wv2EnvReady(g_Env)',
    )
    code = code.replace(
        '    Debug.Print "controller ready in env"' + "\n" + "    RaiseEvent wv2CtrlReady(createdController)",
        '    pluginExample.TraceOhada "wv2Environment.Event:controller ready"' + "\n" + "    RaiseEvent wv2CtrlReady(createdController)",
    )
    return code


def copy_support_files(target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    Path(USERDATA_PATH).mkdir(parents=True, exist_ok=True)
    shutil.copy2(VENDOR_TLB, target_dir / VENDOR_TLB.name)
    shutil.copy2(resolve_loader_dll(), target_dir / VENDOR_DLL.name)
    (target_dir / EDGE_HOST_FILE).write_text(EDGE_HOST_HTML, encoding="utf-8")


def ensure_standard_module(vb_project, name: str):
    try:
        return vb_project.VBComponents(name)
    except Exception:
        component = vb_project.VBComponents.Add(1)
        component.Name = name
        return component


def replace_component_code(component, code: str) -> None:
    module = component.CodeModule
    count = module.CountOfLines
    if count:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def update_webview_reference(vb_project, tlb_path: Path) -> None:
    to_remove = []
    for ref in vb_project.References:
        try:
            if ref.Name == "WebView2_edit":
                to_remove.append(ref)
        except Exception:
            pass

    for ref in to_remove:
        try:
            vb_project.References.Remove(ref)
        except Exception:
            pass

    vb_project.References.AddFromFile(str(tlb_path))


def format_start_sheet(workbook) -> None:
    sheet = workbook.Worksheets(1)
    sheet.Name = "START"
    sheet.Cells.Clear()

    sheet.Range("A1").Value = "OHADA COMPTA"
    sheet.Range("A2").Value = "Version Excel embarquee du site"
    sheet.Range("A4").Value = "Si l'application ne s'ouvre pas automatiquement, activez les macros puis relancez le fichier."
    sheet.Range("A6").Value = "URL chargee dans Excel :"
    sheet.Range("A7").Value = LIVE_SITE_URL
    sheet.Range("A9").Value = "Fichier genere pour afficher le vrai site OHADA-Compta dans Excel via Edge WebView2."

    sheet.Range("A1").Font.Size = 22
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A2").Font.Size = 12
    sheet.Range("A4").Font.Size = 11
    sheet.Range("A7").Font.Color = 0xAA6C00
    sheet.Range("A7").Font.Underline = True
    sheet.Hyperlinks.Add(Anchor=sheet.Range("A7"), Address=LIVE_SITE_URL, TextToDisplay=LIVE_SITE_URL)

    for column in ["A", "B", "C", "D", "E", "F"]:
        sheet.Columns(column).ColumnWidth = 24

    sheet.Range("A1:F20").Interior.Color = 0x0B110A
    sheet.Range("A1:F20").Font.Color = 0xE6E6E6
    sheet.Range("A1:F20").VerticalAlignment = -4160  # xlVAlignTop
    sheet.Range("A1:F20").HorizontalAlignment = -4131  # xlLeft
    sheet.Rows("1:20").RowHeight = 24
    sheet.Range("A1:F20").WrapText = True


def build_workbook(target_workbook: Path) -> None:
    target_workbook.parent.mkdir(parents=True, exist_ok=True)
    copy_support_files(target_workbook.parent)
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

        update_webview_reference(vb_project, target_workbook.parent / VENDOR_TLB.name)

        replace_component_code(vb_project.VBComponents("AppConstants"), patched_app_constants_code())
        replace_component_code(vb_project.VBComponents("APIFunctions"), APIFUNCTIONS_CODE)
        replace_component_code(vb_project.VBComponents("WV2Globals"), patched_wv2_globals_code())
        replace_component_code(vb_project.VBComponents("MemoryFunctions"), patched_memory_functions_code())
        replace_component_code(vb_project.VBComponents("clsWebViewEventHandlers"), patched_cls_webview_event_handlers_code())
        replace_component_code(vb_project.VBComponents("clsWebViewScriptCompleteHandler"), patched_cls_webview_script_complete_handler_code())
        replace_component_code(vb_project.VBComponents("clsWebViewContentHandler"), patched_cls_webview_content_handler_code())
        replace_component_code(vb_project.VBComponents("ThisWorkbook"), THISWORKBOOK_CODE)
        replace_component_code(vb_project.VBComponents("pluginLoader"), PLUGIN_LOADER_CODE)
        replace_component_code(vb_project.VBComponents("pluginExampleCls"), PLUGIN_EXAMPLE_CLS_CODE)
        replace_component_code(vb_project.VBComponents("factory"), FACTORY_CODE)
        replace_component_code(vb_project.VBComponents("wv2"), patched_wv2_code())
        replace_component_code(vb_project.VBComponents("wv2Environment"), patched_wv2_environment_code())
        replace_component_code(vb_project.VBComponents("UserForm1"), USERFORM1_CODE)
        replace_component_code(vb_project.VBComponents("pluginExample"), MOD_OHADA_HOST_CODE)

        format_start_sheet(workbook)

        workbook.Save()
        workbook.Close(SaveChanges=True)
    finally:
        excel.EnableEvents = True
        excel.Quit()
        pythoncom.CoUninitialize()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build an Excel workbook that opens the live OHADA-Compta site inside Excel via WebView2.")
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Primary workbook output path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--desktop-output",
        type=Path,
        default=DEFAULT_DESKTOP_OUTPUT,
        help=f"Desktop workbook output path (default: {DEFAULT_DESKTOP_OUTPUT})",
    )
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
