VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Tricks"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin prjVBTricks.Button btnCancel 
      Height          =   540
      Left            =   5280
      TabIndex        =   7
      Top             =   1980
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   953
      Caption         =   "Cancel"
   End
   Begin prjVBTricks.Button btnOk 
      Default         =   -1  'True
      Height          =   540
      Left            =   5280
      TabIndex        =   6
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   953
      Caption         =   "Ok"
   End
   Begin VB.ListBox lstTricks 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H00000080&
      Height          =   2595
      ItemData        =   "frmMain.frx":7D26E
      Left            =   300
      List            =   "frmMain.frx":7D270
      TabIndex        =   1
      Top             =   1380
      Width           =   4935
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picBGround 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   615
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4620
      Width           =   5940
      Begin VB.PictureBox picMarquee 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5940
         Picture         =   "frmMain.frx":7D272
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   748
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   11220
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   300
      Width           =   6615
   End
   Begin VB.Label lblTricks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tricks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   1140
      Width           =   3195
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   435
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   315
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************'
'*                                                                      *'
'*  Email : AJB2001LG@YAHOO.COM                                         *'
'*                                                                      *'
'*  Date Created :   February 1, 2002                                   *'
'*  Date Finished:   February 3, 2002                                   *'
'*                                                                      *'
'*                                                                      *'
'************************************************************************'

Option Explicit ' all variables must be declared

' just wait for my next version.

'=============================================================================================
Dim OldMarqueeLeft As Integer ' old marquee left
'=============================================================================================

Private Sub Button1_Click()

End Sub

'=============================================================================================
Private Sub Form_Load()
    Dim i As Integer
    
    ' Me.hWnd or  frmMain.hWnd - handle to a window
    ' Me.Picture or frmMain.Picture - picture to be loaded
    ' vbWhite - the color that specifies the transparent area
    ShapedForm.Shape Me.hwnd, Me.Picture, vbWhite
    
    For i = 1 To 11
        lstTricks.AddItem "  " & "Trick #: " & Format$(i, "00")
    Next i
    
    lstTricks.ListIndex = lstTricks.TopIndex
    OldMarqueeLeft = picMarquee.Left ' get old marquee left
    tmrTimer.Enabled = True ' start the timer
End Sub

'=============================================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then DragObject Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'=============================================================================================
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown vbLeftButton, 0, 0, 0
End Sub

'=============================================================================================
Private Sub btnOk_Click()
    Select Case lstTricks.ListIndex
    Case Is = 0
        frmTrick01.Show vbModal
    Case Is = 1
        frmTrick02.Show vbModal
    Case Is = 2
        frmTrick03.Show vbModal
    Case Is = 3
        frmTrick04.Show vbModal
    Case Is = 4
        frmTrick05.Show vbModal
    Case Is = 5
        frmTrick06.Show vbModal
    Case Is = 6
        frmTrick07.Show vbModal
    Case Is = 7
        frmTrick08.Show vbModal
    Case Is = 8
        frmTrick09.Show vbModal
    Case Is = 9
        frmTrick10.Show vbModal
    Case Is = 10
        Me.Hide ' hide the form to perform quick animation
        frmTrick11.Show vbModal
        Me.Show ' show the form
    End Select
End Sub

'=============================================================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lstTricks_DblClick()
    btnOk_Click
End Sub

'=============================================================================================
Private Sub tmrTimer_Timer()
    Static i As Integer
    Dim Title As String
    
    picMarquee.Left = picMarquee.Left - 1
    If Abs(picMarquee.Left) >= picMarquee.Width Then _
        picMarquee.Left = OldMarqueeLeft
    
    Title = "VB Tricks ver 1.0"
    
    If i <> Len(Title) Then
        i = i + 1
        lblTitle(0).Caption = Left(Title, i)
        lblTitle(1).Caption = Left(Title, i)
    End If
End Sub
'=============================================================================================
