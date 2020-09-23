VERSION 5.00
Begin VB.Form frmTrick01 
   Caption         =   "Trick #: 01"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   0
      ScaleHeight     =   179
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Line linHLine 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   80
      End
      Begin VB.Line linVLine 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         X1              =   28
         X2              =   28
         Y1              =   0
         Y2              =   80
      End
   End
End
Attribute VB_Name = "frmTrick01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

'=============================================================================================
Private Sub Form_Load()
    Dim sw As Single, sh As Single
    Dim xmid As Single, ymid As Single
    
    sw = picPicture.ScaleWidth
    sh = picPicture.ScaleHeight
    
    xmid = sw / 2 ' get X midpoint
    ymid = sh / 2 ' get Y midpoint
    
    ' Horizontal line
    linHLine.x1 = 0
    linHLine.y1 = ymid ' move to the center Y-Axis
    linHLine.x2 = sw
    linHLine.y2 = ymid ' move to the center Y-Axis
    
    ' Vertical line
    linVLine.x1 = xmid ' move to the center X-Axis
    linVLine.y1 = 0
    linVLine.x2 = xmid ' move to the center Y-Axis
    linVLine.y2 = sh
End Sub

'=============================================================================================
Private Sub Form_Resize()
    On Error Resume Next ' turns off error handling
    
    picPicture.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'=============================================================================================
Private Sub picPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Horizontal line
    linHLine.x1 = 0
    linHLine.y1 = y ' follow the mouse movement for Y-Axis
    linHLine.x2 = picPicture.ScaleWidth
    linHLine.y2 = y ' follow the mouse movement for Y-Axis
    
    ' Vertical line
    linVLine.x1 = x ' follow the mouse movement for X-Axis
    linVLine.y1 = 0
    linVLine.x2 = x ' follow the mouse movement for X-Axis
    linVLine.y2 = picPicture.ScaleHeight
End Sub
'=============================================================================================
