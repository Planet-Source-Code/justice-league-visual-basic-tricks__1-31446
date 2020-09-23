VERSION 5.00
Begin VB.UserControl ShapedPicture 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   1200
   ScaleWidth      =   1440
   Begin VB.PictureBox picHolder 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "ShapedPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'=============================================================================================
Public Enum BackStyleConstant
    Transparent
    Opaque
End Enum

'=============================================================================================
Dim m_AutoSize As Boolean
'=============================================================================================

'=============================================================================================
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()
'=============================================================================================

'=============================================================================================
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

'=============================================================================================
Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If UserControl.Picture And AutoSize Then
        UserControl.Width = picHolder.Width
        UserControl.Height = picHolder.Height
    End If
End Property

'=============================================================================================
Public Property Get BackStyle() As BackStyleConstant
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

'=============================================================================================
Public Property Let BackStyle(ByVal New_BackStyle As BackStyleConstant)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'=============================================================================================
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'=============================================================================================
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'=============================================================================================
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'=============================================================================================
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'=============================================================================================
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

'=============================================================================================
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'=============================================================================================
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

'=============================================================================================
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'=============================================================================================
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
    Set MaskPicture = UserControl.MaskPicture
End Property

'=============================================================================================
Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property

'=============================================================================================
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

'=============================================================================================
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'=============================================================================================
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

'=============================================================================================
Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    If UserControl.Picture Then Set picHolder.Picture = UserControl.Picture
    AutoSize = AutoSize
End Property

'=============================================================================================
Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

'=============================================================================================
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'=============================================================================================
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'=============================================================================================
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'=============================================================================================
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'=============================================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'=============================================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'=============================================================================================
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'=============================================================================================
Public Sub Refresh()
    UserControl.Refresh
End Sub

'=============================================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoSize = PropBag.ReadProperty("AutoSize", False)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", Opaque)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'=============================================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, False)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, Opaque)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'=============================================================================================
Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Display the copyright dialog."
Attribute ShowAbout.VB_UserMemId = -552
    MsgBox "ShapedPicture ver 1.0" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : AJB2001LG@YAHOO.COM", , "ShapedPicture"
End Sub
'=============================================================================================


