Attribute VB_Name = "Module1"
Option Explicit

'Windows API
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function RegisterShellHook Lib "Shell32" Alias "#181" (ByVal hWnd As Long, ByVal nAction As Long) As Long

'Undocumented (at least that I could find) Windows API.
'  Use with caution
Declare Function RegisterShellHookWindow Lib "user32" (ByVal hWnd As Long) As Long

'Constants for Windows API
Public Const HSHELL_WINDOWCREATED = 1
Public Const HSHELL_WINDOWDESTROYED = 2
Public Const HSHELL_ACTIVATESHELLWINDOW = 3
Public Const HSHELL_WINDOWACTIVATED = 4
Public Const HSHELL_GETMINRECT = 5
Public Const HSHELL_REDRAW = 6
Public Const HSHELL_TASKMAN = 7
Public Const HSHELL_LANGUAGE = 8

Public Const WM_NCDESTROY = &H82

Public Const GWL_WNDPROC = -4

Public Const WH_SHELL = 10
Public Const WH_CBT As Long = 5

Public Const GW_OWNER = 4
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000

Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const RSH_REGISTER_PROGMAN = 2
Public Const RSH_REGISTER_TASKMAN = 3

'Variables
Private lpPrevWndProc As Long ' Address of previos window proc
Private msgShellHook As Long  ' Msg number of "SHELLHOOK" message
Public lHook As Long

'Called repetivly in response to EnumWindows command.  Add window (if
'  visible) to the list of windows.
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) _
  As Long
  
    Form1.WindowCreated hWnd 'Just pretend we saw the window be created
  
  EnumWindowsProc = True
End Function

'Start the subclassing of the window
Public Sub Hook(hWnd As Long)
  lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'Stop the subclasing of the window
Public Sub Unhook(hWnd As Long)
  SetWindowLong hWnd, GWL_WNDPROC, lpPrevWndProc
End Sub

'Main Entry point.  Setup the system wide WH_SHELL hook, and start the
'  subclassing of the form to view messages.
Public Sub StartHook(hWnd As Long)

  'This is the message that Shell32's ShellHookProc sends us whenever
  '  a shell hook occurs
  msgShellHook = RegisterWindowMessage("SHELLHOOK")

  'Load the Shell32 library, and find the ShellHookProc so we can pass
  '  it to SetWindowsHookEx to create the Shell Hook
  Dim hLibShell As Long
  Dim lpHookProc As Long
  
  hLibShell = LoadLibrary("shell32.dll")
  lpHookProc = GetProcAddress(hLibShell, "ShellHookProc")
  'MsgBox hLibShell & " , " & lpHookProc
  'Initialize ShellHookProc
   RegisterShellHookWindow hWnd
  
  SetWindowsHookEx WH_CBT, lpHookProc, hLibShell, 0
   
  'Start the subclassing of the window so we get the "SHELLHOOK" '
  '  messages generated from ShellHookProc
  Hook hWnd

  'Enumurate through the windows so we can get a list of running windows.
  EnumWindows AddressOf EnumWindowsProc, 0

End Sub

'Subclassing procedure, look for "SHELLHOOK" messages and process
Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
  ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case uMsg
    Case WM_NCDESTROY  'If we receive this message, unhook our subclassing
                       '  routing, to prevent crashing the app as it closes
      Unhook hWnd
    Case msgShellHook  'This is the message generated from Shell32's
                       '  ShellHookProc, decode it, and send the results
                       '  up to the From1's handlers
      Select Case wParam
        Case HSHELL_WINDOWCREATED
            Form1.WindowCreated lParam
        Case HSHELL_WINDOWDESTROYED
          Form1.WindowDestroyed lParam
        Case HSHELL_REDRAW
          Form1.WindowRedraw lParam
        Case HSHELL_WINDOWACTIVATED
          Form1.WindowActivated lParam
      End Select
  End Select

  WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
End Function

'
' End of module code
