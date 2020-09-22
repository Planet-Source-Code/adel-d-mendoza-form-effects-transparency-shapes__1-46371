Attribute VB_Name = "Module1"

' To illustrate some form effects: transparency and shapes.  This code allows
' for free switching between normal and transparent states of the same form (and
' free toggling between transparency with frame and without frame), and shows
' shaped forms as well (rectangular form with rounded corners, oval form and round
' form).  In all cases, the code also shows how to drag a form.

Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'RGN_OR creates the union of combined regions
Public Const RGN_OR = 2

'RGN_DIFF creates the intersection of combined regions
Public Const RGN_DIFF = 4
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public xp As Long, yp As Long
Public mShape As Integer

Sub Main()
    If Program_Is_Already_Running Then
        GoTo Noduplicate
    End If
    xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY
    frmTransparentForm.Show
    Exit Sub
    
Noduplicate:
    MsgBox "Program is already running"
End Sub

Function Program_Is_Already_Running()
    Program_Is_Already_Running = False
    If (App.PrevInstance = True) Then
        Program_Is_Already_Running = True
    End If
End Function

Function UnloadIfExist(FormName As String) As Boolean
    On Error Resume Next
    Dim mloaded As Boolean
    mloaded = False
    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
        If UCase(Forms(i).Name) = UCase(FormName) Then
             mloaded = True
             Unload Forms(i)
             Exit For
        End If
    Next
    UnloadIfExist = mloaded
End Function


    
