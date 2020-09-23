VERSION 5.00
Begin VB.Form frmTrick09 
   Caption         =   "Trick #: 09"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton03 
      Caption         =   "Button 3"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton02 
      Caption         =   "Button 2"
      Height          =   495
      Left            =   2220
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton01 
      Caption         =   "Button 1"
      Height          =   495
      Left            =   780
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTrick09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================================
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'=============================================================================================

'=============================================================================================
Const BDR_INNER = &HC
Const BDR_OUTER = &H3
Const BDR_RAISED = &H5
Const BDR_RAISEDINNER = &H4
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKEN = &HA
Const BDR_SUNKENINNER = &H8
Const BDR_SUNKENOUTER = &H2
'=============================================================================================

'=============================================================================================
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'=============================================================================================

'=============================================================================================
Private Sub cmdButton01_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hDC As Long
    Dim rRect As RECT
    
    ' cmdButton01
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton01.Width
    rRect.Bottom = cmdButton01.Height
    
    hDC = GetDC(cmdButton01.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_SUNKEN, &H100F
    
    ' cmdButton02
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton02.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
    
    ' cmdButton03
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton03.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
End Sub

'=============================================================================================
Private Sub cmdButton02_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hDC As Long
    Dim rRect As RECT
    
    ' cmdButton01
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton01.Width
    rRect.Bottom = cmdButton01.Height
    
    hDC = GetDC(cmdButton01.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
    
    ' cmdButton02
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton02.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_SUNKENINNER, &H100F
    
    ' cmdButton03
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton03.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
End Sub

'=============================================================================================
Private Sub cmdButton03_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hDC As Long
    Dim rRect As RECT
    
    ' cmdButton01
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton01.Width
    rRect.Bottom = cmdButton01.Height
    
    hDC = GetDC(cmdButton01.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
    
    ' cmdButton02
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton02.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISED, &H100F
    
    ' cmdButton03
    rRect.Left = 0
    rRect.Top = 0
    rRect.Right = cmdButton02.Width
    rRect.Bottom = cmdButton02.Height
    
    hDC = GetDC(cmdButton03.hwnd) ' get button hdc
    
    DrawEdge hDC, rRect, BDR_RAISEDINNER, &H100F
End Sub
'=============================================================================================
