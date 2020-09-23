VERSION 5.00
Begin VB.Form frmTrick07 
   Caption         =   "Trick #: 07"
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optDots 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Option Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   3915
   End
   Begin VB.PictureBox picShark 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   4080
      ScaleHeight     =   2265
      ScaleWidth      =   4140
      TabIndex        =   2
      Top             =   2700
      Width           =   4140
   End
   Begin VB.TextBox txtDots 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.CommandButton cmdShark 
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   4020
   End
   Begin VB.Menu mnuClickMe 
      Caption         =   "&Click Me"
   End
End
Attribute VB_Name = "frmTrick07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

' shaped.dll

'=============================================================================================
Dim FilePath As String
Dim ShapeObject As New clsShaped
'=============================================================================================

Private Sub Form_Load()
    FilePath = App.Path & "\Pictures\Shark01.bmp"
    
    cmdShark.Picture = LoadPicture(FilePath)
    picShark.Picture = LoadPicture(FilePath)
End Sub

'=============================================================================================
Private Sub mnuClickMe_Click()
    ' just get the hwnd
    ShapeObject.Shape cmdShark.hwnd, LoadPicture(FilePath), vbWhite
    ShapeObject.Shape txtDots.hwnd, LoadPicture(FilePath), vbWhite
    ShapeObject.Shape picShark.hwnd, LoadPicture(FilePath), vbWhite
    ShapeObject.Shape optDots.hwnd, LoadPicture(FilePath), vbWhite
End Sub
'=============================================================================================
