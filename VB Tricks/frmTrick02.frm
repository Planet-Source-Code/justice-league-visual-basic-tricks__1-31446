VERSION 5.00
Begin VB.Form frmTrick02 
   Caption         =   "Trick #: 02"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDrawing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   2925
      TabIndex        =   0
      Top             =   0
      Width           =   2955
   End
   Begin VB.Menu mnuIncrease 
      Caption         =   "&Increase"
   End
   Begin VB.Menu mnuDecrease 
      Caption         =   "&Decrease"
   End
End
Attribute VB_Name = "frmTrick02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

'=============================================================================================
Dim OldSW As Single ' ScaleWidth
Dim OldSH As Single ' ScaleHeight
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    OldSW = picDrawing.ScaleWidth  ' get old scalewidth
    OldSH = picDrawing.ScaleHeight ' get old scaleheight
    
    Draw
End Sub

'=============================================================================================
Private Sub Draw()
    Dim sw As Single, sh As Single
    Dim xmid As Single, ymid As Single
    
    sw = picDrawing.ScaleWidth  ' get scalewidth
    sh = picDrawing.ScaleHeight ' get scaleheight
    xmid = sw / 2 ' get X midpoint
    ymid = sh / 2 ' get Y midpoint
    
    picDrawing.Cls ' removes any graphics drawn
    
    ' just draw anything
    picDrawing.Circle (xmid, ymid), 500
    picDrawing.Circle (xmid, ymid), 100
    picDrawing.Line (xmid - 600, ymid - 600)-(xmid + 600, ymid + 600)
    picDrawing.Line (0, 0)-(xmid, ymid), , B
End Sub

'=============================================================================================
' you may change the width and height of the picture,
' but the scalewidth and scaleheight
' must be fixed in order to zoom the picture
Private Sub mnuDecrease_Click()
    On Error Resume Next ' turns off error handling
    
    picDrawing.Width = picDrawing.Width * 0.75    ' Decrease by 25%
    picDrawing.Height = picDrawing.Height * 0.75  ' Decrease by 25%
    picDrawing.ScaleWidth = OldSW  ' restore old scalewidth
    picDrawing.ScaleHeight = OldSH ' restore old scaleheight
    Draw
End Sub

'=============================================================================================
Private Sub mnuIncrease_Click()
    On Error Resume Next ' turns off error handling
    
    picDrawing.Width = picDrawing.Width / 0.75    ' Increase by 25%
    picDrawing.Height = picDrawing.Height / 0.75  ' Increase by 25%
    picDrawing.ScaleWidth = OldSW  ' restore old scalewidth
    picDrawing.ScaleHeight = OldSH ' restore old scaleheight
    Draw
End Sub
'=============================================================================================
