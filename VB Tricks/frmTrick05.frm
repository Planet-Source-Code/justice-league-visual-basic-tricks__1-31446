VERSION 5.00
Begin VB.Form frmTrick05 
   BorderStyle     =   0  'None
   Caption         =   "Trick #: 05"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1140
      Width           =   675
   End
End
Attribute VB_Name = "frmTrick05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

' shaped.dll

'=============================================================================================
Dim ShapeForm As New clsShaped
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    Dim FilePath As String
    
    FilePath = App.Path & "\PICTURES\SHARK01.BMP"
    
    Me.Picture = LoadPicture(FilePath)
    ' Me.hWnd or  frmTrick05.hWnd - handle to a window
    ' Me.Picture or frmTrick05.Picture - picture to be loaded
    ' vbWhite - the color that specifies the transparent area
    ShapeForm.Shape Me.hwnd, Me.Picture, vbWhite
End Sub

'=============================================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' drag the form
    If Button And vbLeftButton Then DragObject Me.hwnd
End Sub
'=============================================================================================
'=============================================================================================
Private Sub cmdExit_Click()
    Me.Hide
End Sub


