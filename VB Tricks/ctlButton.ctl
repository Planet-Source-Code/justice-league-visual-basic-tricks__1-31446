VERSION 5.00
Begin VB.UserControl Button 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   FillColor       =   &H80000005&
   MaskColor       =   &H0000FFBF&
   PropertyPages   =   "ctlButton.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   1140
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'=============================================================================================
Const BTN_UP = 101
Const BTN_DOWN = 102
Const BTN_FOCUS = 103
Const BTN_DISABLED = 104
'=============================================================================================

'=============================================================================================
Const MIN_WIDTH = 1140
Const MIN_HEIGHT = 540
'=============================================================================================

'=============================================================================================
Const SRCCOPY = &HCC0020
Const SRCINVERT = &H660046
'=============================================================================================

'=============================================================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'=============================================================================================

'=============================================================================================
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'=============================================================================================

'=============================================================================================
Dim IsButtonSelect As Boolean
'=============================================================================================

'=============================================================================================
Private Function GetPicture(ByVal opt As Integer) As Picture
    Dim i As Integer, w As Integer
    
    Set picMask.Picture = LoadPicture()
    Set picSrc.Picture = LoadResPicture(opt, vbResBitmap)
    
    picDest.Width = UserControl.Width
    UserControl.Height = picDest.Height
    
    w = picDest.Width \ Screen.TwipsPerPixelX - 22
    
    If opt <> BTN_DOWN Then
        lblCaption.Move 150, 60, UserControl.Width - Screen.TwipsPerPixelX * 22, 375
        lblShadow.Move 165, 75, UserControl.Width - Screen.TwipsPerPixelX * 22, 375
        Call BitBlt(picDest.hDC, 0, 0, 20, picDest.Height, picSrc.hDC, 0, 0, SRCCOPY)
        Call BitBlt(picDest.hDC, w, 0, 22, picDest.Height, picSrc.hDC, 21, 0, SRCCOPY)
        
         For i = 20 To w
            Call BitBlt(picDest.hDC, i, 0, 1, picDest.Height, picSrc.hDC, 20, 0, SRCCOPY)
        Next i
    Else
        lblCaption.Move 180, 90, UserControl.Width - Screen.TwipsPerPixelX * 22, 375
        lblShadow.Move 195, 105, UserControl.Width - Screen.TwipsPerPixelX * 22, 375
        Call BitBlt(picDest.hDC, 0, 0, 22, picDest.Height, picSrc.hDC, 0, 0, SRCCOPY)
        Call BitBlt(picDest.hDC, w + 1, 0, 23, picDest.Height, picSrc.hDC, 22, 0, SRCCOPY)

        For i = 22 To w + 1
            Call BitBlt(picDest.hDC, i, 0, 1, picDest.Height, picSrc.hDC, 22, 0, SRCCOPY)
        Next i
    End If
    
    Set picDest.Picture = picDest.Image
    picMask.Width = picDest.Width
    picMask.Height = picDest.Height
        
    Call BitBlt(picMask.hDC, 0, 0, picMask.Width, picMask.Height, picDest.hDC, 0, 0, SRCINVERT)
        
    Set picMask.Picture = picMask.Image
    Set UserControl.MaskPicture = picMask.Picture
    Set GetPicture = picDest.Picture
End Function

'=============================================================================================
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
        
    lblShadow.Caption = lblCaption.Caption
End Property

'=============================================================================================
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    UserControl.Picture = IIf(Enabled, GetPicture(BTN_UP), GetPicture(BTN_DISABLED))
End Property

'=============================================================================================
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'=============================================================================================
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'=============================================================================================
Private Sub lblCaption_Click()
    RaiseEvent Click
End Sub

'=============================================================================================
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Then lblCaption_Click
End Sub

'=============================================================================================
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    On Error Resume Next
    
    RaiseEvent KeyDown(KeyCode, Shift)
        
    Select Case KeyCode
    Case Is = vbKeyReturn
        If Shift = 0 Then lblCaption_Click
    Case Is = vbKeySpace
        If Shift = 0 Then
            If Not IsButtonSelect Then
                lblCaption_MouseDown vbLeftButton, 0, 0, 0
                IsButtonSelect = True
            End If
        End If
    Case vbKeyRight, vbKeyDown
        SendKeys "{Tab}"
    Case vbKeyLeft, vbKeyUp
        SendKeys "+{Tab}"
    Case Else
        If IsButtonSelect Then
            lblCaption_MouseUp vbLeftButton, 0, 0, 0
            IsButtonSelect = False
            lblCaption_Click
        End If
    End Select
End Sub

'=============================================================================================
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'=============================================================================================
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    
    If IsButtonSelect Then
        lblCaption_MouseUp vbLeftButton, 0, 0, 0
        IsButtonSelect = False
        lblCaption_Click
    End If
End Sub

'=============================================================================================
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, _
        ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
        
    If Button And vbLeftButton Then UserControl.Picture = GetPicture(BTN_DOWN)
End Sub

'=============================================================================================
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    
    Static flutterUp As Boolean
    Static flutterDn As Boolean
    
    RaiseEvent MouseMove(Button, Shift, _
        ScaleX(X + lblCaption.Left, vbTwips, vbContainerPosition), _
        ScaleY(Y + lblCaption.Height, vbTwips, vbContainerPosition))
        
    If Button And vbLeftButton Then
        If X < 0 Or X > lblCaption.Width Or Y < 0 Or Y > lblCaption.Height Then
            If Not flutterUp Then
                flutterUp = True
                UserControl_EnterFocus
            End If
            flutterDn = False
        Else
            If Not flutterDn Then
                flutterDn = True
                UserControl.Picture = GetPicture(BTN_DOWN)
            End If
            flutterUp = False
        End If
    End If
    
    UserControl.SetFocus
End Sub

'=============================================================================================
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    
    RaiseEvent MouseUp(Button, Shift, _
        ScaleX(X + lblCaption.Left, vbTwips, vbContainerPosition), _
        ScaleY(Y + lblCaption.Height, vbTwips, vbContainerPosition))

    If Button And vbLeftButton Then UserControl_EnterFocus
End Sub

'=============================================================================================
Private Sub UserControl_EnterFocus()
    UserControl.Picture = GetPicture(BTN_FOCUS)
    lblCaption.FontSize = 10
    lblShadow.FontSize = 10
    lblCaption.Top = lblCaption.Top + Screen.TwipsPerPixelX * 4
    lblShadow.Top = lblShadow.Top + Screen.TwipsPerPixelX * 4
End Sub

'=============================================================================================
Private Sub UserControl_ExitFocus()
    UserControl.Picture = GetPicture(BTN_UP)
    lblCaption.FontSize = 12
    lblShadow.FontSize = 12
End Sub

'=============================================================================================
Private Sub UserControl_Resize()
    If MIN_WIDTH > UserControl.Width Then UserControl.Width = MIN_WIDTH
    If MIN_HEIGHT <> UserControl.Height Then UserControl.Height = MIN_HEIGHT
    
    UserControl.Picture = IIf(Enabled, GetPicture(BTN_UP), GetPicture(BTN_DISABLED))
End Sub

'=============================================================================================
Private Sub UserControl_Show()
    lblShadow.Caption = lblCaption.Caption
    
    UserControl.Picture = IIf(Enabled, GetPicture(BTN_UP), GetPicture(BTN_DISABLED))
End Sub

'=============================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

'=============================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

'=============================================================================================
Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Display the copyright dialog."
Attribute ShowAbout.VB_UserMemId = -552
    MsgBox "Button ver 1.0" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : AJB2001LG@YAHOO.COM", , "Button"
End Sub
'=============================================================================================
