VERSION 5.00
Begin VB.Form frmTrick03 
   Caption         =   "Trick #: 03"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   2925
      TabIndex        =   3
      Top             =   1980
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "Printer"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Clipboard"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
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
End
Attribute VB_Name = "frmTrick03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

Private Sub Form_Load()
    Draw
End Sub

Private Sub cmdClipboard_Click()
    Set picPicture.Picture = picDrawing.Image ' convert image to picture
    Clipboard.Clear ' cleans out the contents  of the windows clipboard
    Clipboard.SetData picPicture.Picture ' save the picture to the clipboard
    picPicture.Picture = LoadPicture("") ' remove picture
End Sub

Private Sub cmdPrinter_Click()
    Dim xpos As Single, ypos As Single
    On Error Resume Next ' turns off error handling
    
    Set picPicture.Picture = picDrawing.Image ' convert image to picture
    Printer.ScaleMode = vbTwips
    xpos = (Printer.ScaleWidth - picPicture.ScaleWidth) / 2
    ypos = (Printer.ScaleHeight - picPicture.ScaleHeight) / 2
    Printer.PaintPicture picPicture.Picture, xpos, ypos ' draws the contents of a graphics on a printer object
    Printer.EndDoc ' start printing...
    picPicture.Picture = LoadPicture("") ' remove picture
End Sub

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
    picDrawing.Line (0, 0)-(xmid, ymid), , B
End Sub

