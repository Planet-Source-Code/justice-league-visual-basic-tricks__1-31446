VERSION 5.00
Begin VB.Form frmTrick04 
   AutoRedraw      =   -1  'True
   Caption         =   "Trick #: 04"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuScreenShot 
      Caption         =   "&Screen Shot"
   End
End
Attribute VB_Name = "frmTrick04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

'=============================================================================================
Const SRCCOPY = &HCC0020
'=============================================================================================

'=============================================================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'=============================================================================================

'=============================================================================================
Private Sub mnuScreenShot_Click()
    Dim hDC As Long
    
    Me.Cls ' removes any graphics drawn
    hDC = GetDC(GetDesktopWindow()) ' get desktop hDC
    
    Call BitBlt(Me.hDC, 0, 0, Screen.Width, Screen.Height, hDC, 0, 0, SRCCOPY)
End Sub
'=============================================================================================
