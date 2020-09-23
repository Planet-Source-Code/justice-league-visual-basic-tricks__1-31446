VERSION 5.00
Begin VB.Form frmTrick06 
   AutoRedraw      =   -1  'True
   Caption         =   "Trick #: 06"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjVBTricks.ShapedPicture shpPicture 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
   End
   Begin VB.Shape shpRectangle 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillStyle       =   6  'Cross
      Height          =   1275
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmTrick06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

' see shark01.bmp
' see shark02.bmp - inverted

'=============================================================================================
Private Sub Form_Load()
    Dim FilePath As String
    
    FilePath = App.Path & "\Pictures" ' pictures path
    
    shpPicture.AutoSize = True
    shpPicture.BackStyle = Transparent
    shpPicture.MaskColor = vbBlack ' the color that specifies the transparent area
    Set shpPicture.Picture = LoadPicture(FilePath & "\SHARK01.BMP")
    Set shpPicture.MaskPicture = LoadPicture(FilePath & "\SHARK02.BMP")
End Sub

'=============================================================================================
Private Sub Form_Resize()
    On Error Resume Next ' turns off error handling
    
    Me.Refresh
    shpRectangle.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'=============================================================================================
Private Sub shpPicture_DblClick()
    MsgBox "This is a test"
End Sub
'=============================================================================================
Private Sub shpPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' drag the shpPicture
    If Button And vbLeftButton Then DragObject shpPicture.hwnd
End Sub

Private Sub ShapedPicture1_Click()

End Sub
