VERSION 5.00
Begin VB.Form frmTrick10 
   Caption         =   "Trick #: 10"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3300
      Top             =   1440
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate"
      Height          =   375
      Left            =   5220
      TabIndex        =   2
      Top             =   2940
      Width           =   1275
   End
   Begin VB.PictureBox picMarquee 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3960
      ScaleHeight     =   2745
      ScaleWidth      =   3765
      TabIndex        =   1
      Top             =   60
      Width           =   3795
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   3765
      TabIndex        =   0
      Top             =   60
      Width           =   3795
   End
End
Attribute VB_Name = "frmTrick10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================================
Const HORIZ_LEFT_TEXT = 0
Const HORIZ_RIGHT_TEXT = 1
Const HORIZ_CENTER_TEXT = 2

Const VERT_TOP_TEXT = 4
Const VERT_BOTTOM_TEXT = 8
Const VERT_CENTER_TEXT = 16
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    PutText picText, "Center-Top", , , vbRed, , , HORIZ_CENTER_TEXT Or VERT_TOP_TEXT
    PutText picText, "Center-Bottom", , , vbRed, , , HORIZ_CENTER_TEXT Or VERT_BOTTOM_TEXT
    PutText picText, "Left-Center", , , vbBlue, , , HORIZ_LEFT_TEXT Or VERT_CENTER_TEXT
    PutText picText, "Right-Center", , , vbBlue, , , HORIZ_RIGHT_TEXT Or VERT_CENTER_TEXT
    PutText picText, "Center-Center", , , vbMagenta, , , HORIZ_CENTER_TEXT Or VERT_CENTER_TEXT
    PutText picText, "Customize", 100, 100, , True
    PutText picText, "Customize-Center-Center", 500, 500, vbCyan, True, , HORIZ_CENTER_TEXT Or VERT_CENTER_TEXT
End Sub

'=============================================================================================
Public Sub PutText(ByRef obj As Object, ByVal s As String, Optional ByVal x As Single = 0, Optional ByVal y As Single = 0, _
    Optional ByVal Color As OLE_COLOR = &H0, Optional ByVal IsBold As Boolean = False, Optional ByVal IsItalic As Boolean = False, Optional ByVal Alignment As Byte = (HORIZ_LEFT_TEXT Or VERT_TOP_TEXT))

    Dim w As Single, h As Single
    Dim tw As Single, th As Single
    Dim xpos As Single, ypos As Single
    
    w = obj.ScaleWidth
    h = obj.ScaleHeight
    
    Select Case Alignment
    Case Is = 4     ' (0 4)  HORIZ_LEFT_TEXT OR VERT_TOP_TEXT
        xpos = x
        ypos = y
    Case Is = 5     ' (1 4)  HORIZ_RIGHT_TEXT OR VERT_TOP_TEXT
        xpos = w - obj.TextWidth(s) + x
        ypos = y
    Case Is = 6     ' (2 4)  HORIZ_CENTER_TEXT OR VERT_TOP_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + x
        ypos = y
    Case Is = 8     ' (0 8)  HORIZ_LEFT_TEXT OR VERT_BOTTOM_TEXT
        xpos = x
        ypos = h - obj.TextHeight(s) + y
    Case Is = 9     ' (1 8)  HORIZ_RIGHT_TEXT OR VERT_BOTTOM_TEXT
        xpos = w - obj.TextWidth(s) + x
        ypos = h - obj.TextHeight(s) + y
    Case Is = 10    ' (2 8)  HORIZ_CENTER_TEXT OR VERT_BOTTOM_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + x
        ypos = h - obj.TextHeight(s) + y
    Case Is = 16    ' (0 16) HORIZ_LEFT_TEXT OR VERT_CENTER_TEXT
        xpos = x
        ypos = (h - obj.TextHeight(s)) / 2 + y
    Case Is = 17    ' (1 16) HORIZ_RIGHT_EXT OR VERT_CENTER_TEXT
        xpos = w - obj.TextWidth(s) + x
        ypos = (h - obj.TextHeight(s)) / 2 + y
    Case Is = 18    ' (2 16) HORIZ_CENTER_TEXT OR VERT_CENTER_TEXT
        xpos = (w - obj.TextWidth(s)) / 2 + x
        ypos = (h - obj.TextHeight(s)) / 2 + y
    Case Else
        xpos = x
        ypos = y
    End Select

    obj.ForeColor = Color
    obj.FontBold = IsBold
    obj.FontItalic = IsItalic
    obj.CurrentX = xpos
    obj.CurrentY = ypos
    obj.Print s
End Sub

'=============================================================================================
Private Sub cmdAnimate_Click()
    tmrTimer.Enabled = True
End Sub

'=============================================================================================
Private Sub tmrTimer_Timer()
    Dim s As String
    Static i As Integer
    
    s = "JUST USE YOUR IMAGINATION!!!"
    
    If i = Len(s) Then
        i = 0
        tmrTimer.Enabled = False
        Exit Sub
    End If
    
    i = i + 1
    
    picMarquee.Cls
    PutText picMarquee, Mid$(s, 1, i), , , , True, , HORIZ_CENTER_TEXT Or VERT_CENTER_TEXT
    PutText picMarquee, Mid$(s, 1, i), , , , True, , HORIZ_RIGHT_TEXT Or VERT_TOP_TEXT
    PutText picMarquee, Mid$(s, 1, i), , , , True, , HORIZ_LEFT_TEXT Or VERT_BOTTOM_TEXT
End Sub
'=============================================================================================
