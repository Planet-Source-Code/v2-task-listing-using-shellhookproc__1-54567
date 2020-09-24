VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   7665
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hWndOldActive As Long

'Called by the module anytime another window gains focus.  Since we only
'  know the hwnd of the new window, we need to keep track of the last
'  window to keep focus was (and don't assume that '** ' was put in the
'  caption by us (a window can have '** ' in it's caption))
Public Sub WindowActivated(hWnd As Long)
  Dim i As Integer

  'First off, go through and find the old active window, and remove
  '  the '** ' from the front of the title
  For i = 0 To List1.ListCount - 1
    If List1.ItemData(i) = hWndOldActive Then
      If Mid(List1.List(i), 1, 3) = "** " Then
        List1.List(i) = Mid(List1.List(i), 4)
      End If
      Exit For
    End If
  Next
  
  'Then find the window that was activated, and put '** ' in front of the
  '  caption
  For i = 0 To List1.ListCount - 1
    If List1.ItemData(i) = hWnd Then
      List1.List(i) = "** " & List1.List(i)
      Exit For
    End If
  Next
  
  'Finally, set our variable of the active hwnd
  hWndOldActive = hWnd
  
End Sub

'Called by the module whenever a window caption is changed (or atleast,
'  believed to be changed)
Public Sub WindowRedraw(hWnd As Long)
  Dim strCaption As String
  Dim i As Integer
  strCaption = String(255, " ")
  
  GetWindowText hWnd, strCaption, 254  'Grab the new caption, find the
                                       '  spot in the listbox, and put
                                       '  it in.
  For i = 0 To List1.ListCount - 1
    If List1.ItemData(i) = hWnd Then
      List1.List(i) = strCaption
      List1.ListIndex = i
      Exit For
    End If
  Next
End Sub

'Called by the module whenever a window is created
Public Sub WindowCreated(hWnd As Long)
  Dim i As Integer
  Dim lExStyle    As Long
  Dim bNoOwner    As Boolean
  Dim lreturn     As Long
  Dim sWindowText As String
  If Not hWnd = Me.hWnd Then
    If IsWindowVisible(hWnd) Then
        If GetParent(hWnd) = 0 Then
            bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                
                sWindowText = Space$(1024)
                lreturn = GetWindowText(hWnd, sWindowText, Len(sWindowText))
                If lreturn Then
                   sWindowText = Left$(sWindowText, lreturn)
                     List1.AddItem sWindowText
                     List1.ItemData(List1.NewIndex) = hWnd
                End If
            End If
        End If
    End If
  End If
  
  
End Sub

'Called by the module whenever a window is destroyed
Public Sub WindowDestroyed(hWnd As Long)
  Dim i As Integer
  
  For i = 0 To List1.ListCount - 1 'Loop around looking for the hwnd and
                                   '  remove it from the list
    If List1.ItemData(i) = hWnd Then
      List1.RemoveItem i
      Exit For
    End If
  Next

End Sub

'On form load, start the subclassing of the window,
'  as well as the shell hook
Private Sub Form_Load()
  StartHook Me.hWnd
End Sub

'On form exit, stop the subclassing
Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

' End of form code

