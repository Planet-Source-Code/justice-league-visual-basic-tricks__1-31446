VERSION 5.00
Begin VB.Form frmTrick11 
   Caption         =   "Trick #: 11"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate 5"
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate 4"
      Height          =   495
      Index           =   3
      Left            =   1260
      TabIndex        =   3
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate 3"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate 2"
      Height          =   495
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Animate 1"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmTrick11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variable must be declared

Private Sub cmdAnimate_Click(Index As Integer)
    MoveForm Index, 120
End Sub

Private Sub MoveForm(opt As Integer, Speed As Integer)
    Dim OldFormLeft As Single
    Dim OldFormTop As Single
    
    OldFormLeft = Me.Left
    OldFormTop = Me.Top
    
    If (Me.WindowState <> vbMinimized) And (Me.WindowState <> vbMaximized) Then
        Select Case opt
        Case Is = 0
            Do While (Me.Left >= -Me.Width)
                Me.Left = Me.Left - Speed
                DoEvents
            Loop
            
            Me.Left = OldFormLeft
        Case Is = 1
            Do While (Me.Left <= Screen.Width)
                Me.Left = Me.Left + Speed
                DoEvents
            Loop
            
            Me.Left = OldFormLeft
        Case Is = 2
            Do While (Me.Top >= -Me.Height)
                Me.Top = Me.Top - Speed
                DoEvents
            Loop
            
            Me.Top = OldFormTop
        Case Is = 3
            Do While (Me.Top <= Screen.Height)
                Me.Top = Me.Top + Speed
                DoEvents
            Loop
            
            Me.Top = OldFormTop
        Case Is = 4
            Do While (Me.Left > 0)
                Me.Left = Me.Left - Speed
                DoEvents
            Loop
            
            Do While (Me.Left <= (Screen.Width - Me.Width))
                Me.Left = Me.Left + Speed
                DoEvents
            Loop
            
            Do While (Me.Left >= OldFormLeft)
                Me.Left = Me.Left - Speed
                DoEvents
            Loop
            
            Me.Left = OldFormLeft
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const Speed As Integer = 120
            
    If (Me.WindowState <> vbMinimized) And (Me.WindowState <> vbMaximized) Then
        Do While (Me.Left >= -Me.Width)
            Me.Left = Me.Left - Speed
            Me.Top = Me.Top - Speed
            DoEvents
        Loop
    End If
End Sub
