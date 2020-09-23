VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Out"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1733
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Then press Enter"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblTries 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You have 3 tries left"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the password:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Made by Edward Catchpole
'for Planet Source Code
'Please vote/comment if
'you like this

'you should change the password here:
Const Password = "password"

Dim Tries       As Integer
Dim Term        As Boolean

Private Sub cmdOK_Click()

    Dim sPass       As String
    
    sPass = txtPassword.Text
    
    If LCase(sPass) = LCase(Password) Then
        Call AddToStartUp
        Term = True
        Unload Me
        Set frmMain = Nothing
        End
    Else
        Tries = Tries - 1
        If Tries = 0 Then
            Call AddToStartUp
            MsgBox "This will beep instead of shutdown for your convenience. Remove this line, the beep, the one above the beep and uncomment the ExitWindowsEx statement below", vbInformation, "Lock Out" 'remove this line
            'Call ExitWindowsEx(EWX_FORCE, 0)'uncomment this lie
            Call CursorVisibility(True) 'remove this linem
            Beep 'remove this line
            Unload Me
            Set frmMain = Nothing
            End
        End If
        txtPassword.Text = ""
        txtPassword.SetFocus
        lblTries = "You have " & Tries & " tries left"
    End If
    
End Sub

Private Sub Form_Load()
    
    Tries = 3
    
    lblTries = "You have " & Tries & " tries left"
    
    Me.Show
    
    DoEvents
    
    txtPassword.SetFocus
    'the next command will fool the system into
    'thinking the screensaver is running therefore
    'Windows disables Ctrl+Alt+Delete and Alt+Tab
    SystemParametersInfo SPI_SCREENSAVERRUNNING, 1, 0&, 0&
    
    Call CursorVisibility(False)
    'bring the window to the front
    'permanently
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Call CursorPos 'guess what
    Call SetActiveWindow(hwnd)
    Call AddToStartUp
    Call StartLoop
    
End Sub

Public Function SetActive()

    Call SetActiveWindow(hwnd)
    Call SetForegroundWindow(hwnd)
    
End Function

Public Function AddToStartUp()

    Call SetValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Lock-Out", App.Path & App.EXEName & ".exe")
    
End Function

Public Sub StartLoop()

    Dim x       As Long
    
    Do While x < 1
        Call SetActive
        Call CursorPos
        DoEvents
    Loop
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Term = False Then Cancel = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Term = False Then
        Cancel = True
        Exit Sub
    End If
    
    'the 'screensaver' has finished
    SystemParametersInfo SPI_SCREENSAVERRUNNING, 0, 0&, 0&
    
    If Term = False Then
        'open another copy.
        Call ShellExecute(hwnd, "open", App.Path & App.EXEName & ".exe", "", App.Path, SW_NORMAL)
    End If
    
    Call CursorVisibility(True)
    Unload Me
    Set frmMain = Nothing
    End
    
End Sub
