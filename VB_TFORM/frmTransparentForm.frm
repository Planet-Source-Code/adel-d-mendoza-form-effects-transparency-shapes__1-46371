VERSION 5.00
Begin VB.Form frmTransparentForm 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdShape 
      Caption         =   "Shape 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdShape 
      Caption         =   "Shape 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdShape 
      Caption         =   "Shape 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdFrame 
      BackColor       =   &H80000010&
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CommandButton cmdTransparent 
      Caption         =   "Transparent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdTitle 
      BackColor       =   &H00000000&
      Caption         =   "Transparent Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00404000&
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "frmTransparentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim formEffectIndex As Integer
Dim mFormRegion As Long
Dim msg1 As String, msg2 As String

Private Sub Form_Load()
   Me.ScaleMode = vbPixels
   
   msg1 = "This is a normal form." & vbCrLf & vbCrLf & _
       "Use Transparent command button to toggle" & vbCrLf & _
       "between normal and transparent form" & vbCrLf & _
       "(When you are with the transparent form, you can" & vbCrLf & _
       "use Frame command button to toggle display of the frame of form)." & vbCrLf & vbCrLf & _
       "You can drag the form on any part of the form."
       
   msg2 = "Now you only see the header, this textbox" & vbCrLf & _
       "and command buttons; so you can tell that the" & vbCrLf & _
       "form is transparent.  (If you still want to make" & vbCrLf & _
       "sure, drag the header - the only place to drag" & vbCrLf & _
       "this transparent form if without frame)" & vbCrLf & vbCrLf & _
       "Try Frame and Shape buttons to toggle frame" & vbCrLf & _
       "and to display shaped forms respectively."
       
   formEffectIndex = 0
   Text1.Text = msg1
   cmdTitle.Visible = False
   cmdFrame.Enabled = False
    
   'cmdFrame.Enabled = False
   changeFormEffect formEffectIndex
End Sub
    
Private Sub changeFormEffect(inEffect As Integer)
   Dim w As Single, h As Single
   Dim edge As Single, topEdge As Single
   Dim mLeft, mTop
   Dim i As Integer
   Dim r As Long
   Dim outer As Long, inner As Long
   
   ' Put width/height in same denomination of scalewidth/scaleheight
   w = ScaleX(Width, vbTwips, vbPixels)
   h = ScaleY(Height, vbTwips, vbPixels)
    
   If inEffect = 0 Then
      mFormRegion = CreateRectRgn(0, 0, w, h)
      SetWindowRgn hwnd, mFormRegion, True
      Exit Sub
   End If
    
   mFormRegion = CreateRectRgn(0, 0, 0, 0)
   ' Frame edges measurement
   edge = (w - ScaleWidth) / 2
   topEdge = h - edge - ScaleHeight
   ' Get frame
   If inEffect = 1 Then
      outer = CreateRectRgn(0, 0, w, h)
      inner = CreateRectRgn(edge, topEdge, w - edge, h - edge)
      CombineRgn mFormRegion, outer, inner, RGN_DIFF
   End If
   
   ' Combine regions of controls on form
   For i = 0 To Me.Controls.Count - 1
       If Me.Controls(i).Visible = True Then
          mLeft = ScaleX(Me.Controls(i).Left, Me.ScaleMode, vbPixels) + edge
          mTop = ScaleX(Me.Controls(i).Top, Me.ScaleMode, vbPixels) + topEdge
          r = CreateRectRgn(mLeft, mTop, _
              mLeft + ScaleX(Me.Controls(i).Width, Me.ScaleMode, vbPixels), _
              mTop + ScaleY(Me.Controls(i).Height, Me.ScaleMode, vbPixels))
          CombineRgn mFormRegion, r, mFormRegion, RGN_OR
       End If
   Next
   ' We allow toggle
   SetWindowRgn hwnd, mFormRegion, True
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdTransparent_click()
   If formEffectIndex <> 0 Then
      formEffectIndex = 0
      Text1.Text = msg1
      cmdTitle.Visible = False
      cmdFrame.Enabled = False
   Else
      formEffectIndex = 1
      Text1.Text = msg2
      cmdTitle.Visible = True
      cmdFrame.Enabled = True
   End If
   changeFormEffect formEffectIndex
End Sub

Private Sub cmdFrame_Click()
   If formEffectIndex = 1 Then
      formEffectIndex = 2
   Else
      formEffectIndex = 1
   End If
   changeFormEffect formEffectIndex
End Sub

Private Sub cmdShape_Click(Index As Integer)
   mShape = Index
   UnloadIfExist "frmShapedForm"
   frmShapedForm.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SetWindowRgn hwnd, 0, False
   DeleteObject mFormRegion
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> vbLeftButton Then
      Exit Sub
   End If
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'Unlike frmShapedForm, since frmTransparent is transparent, we have to
'provide a place for user to drag if without frame, so cmdTitle is used.
'-----------------------------------------------------------------------
Private Sub cmdTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> vbLeftButton Then
      Exit Sub
   End If
   ReleaseCapture
   SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



