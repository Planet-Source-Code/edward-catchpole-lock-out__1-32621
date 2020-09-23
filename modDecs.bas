Attribute VB_Name = "modDecs"
Option Explicit

Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4 'force shutdown

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_GETWORKAREA = 48 'get the work area
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1 'bring to top and stay there
Public Const SWP_NOMOVE = &H2 'don't move window
Public Const SWP_NOSIZE = &H1 'don't size window

Public Function CursorPos()

    Call SetCursorPos(0, 0)
    
End Function

Public Function CursorVisibility(Vis As Boolean)

    Call ShowCursor(Vis)
    
End Function
