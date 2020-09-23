VERSION 5.00
Begin VB.Form frmTrick08 
   Caption         =   "Trick #: 08"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2940
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2940
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.ListBox lstRecords 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "frmTrick08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

Private Type FILE_RECORD
    LN As String * 50 ' Lastname (width = 50 chars.)
    FN As String * 50 ' Firstname (width = 50 chars.)
    MI As String * 50 ' Middlename (width = 50 chars.)
End Type

Const FILE_SIGNATURE = "MySignature" ' this will identify the file

'=============================================================================================
Dim FileName As String
'=============================================================================================

'=============================================================================================
Private Sub Form_Load()
    FileName = App.Path & "\temp.txt" ' try to use other files
End Sub

'=============================================================================================
Private Sub cmdSave_Click()
    Dim obj As Object
    Dim cltFR As New Collection
    
    ' first record
    Set obj = New clsRECORDS
    obj.LN = "Einstein"
    obj.FN = "Albert"
    obj.MI = "XXXXXX"
    cltFR.Add obj
    
    ' second record
    Set obj = New clsRECORDS
    obj.LN = "Newton"
    obj.FN = "Isaac"
    obj.MI = "XXXXXX"
    cltFR.Add obj
    
    FileSave FileName, cltFR
    cmdOpen.Enabled = True
End Sub

'=============================================================================================
Private Sub cmdOpen_Click()
    Dim obj As Object
    Dim cltFR As New Collection

    Set obj = New clsRECORDS
    Set cltFR = FileOpen(FileName)
    
    For Each obj In cltFR
        lstRecords.AddItem Trim$(obj.LN) & ", " _
            & Trim$(obj.FN) & " " & Trim$(obj.MI) ' display all records
    Next obj
End Sub

'=============================================================================================
Public Function FileOpen(FileName As String) As Collection
    Dim obj As Object
    Dim InFile As Integer
    Dim Signature As String
    Dim FR As FILE_RECORD
    On Error Resume Next ' turns off error handling
    
    If Dir$(FileName) <> vbNullString Then
        InFile = FreeFile
        Set FileOpen = New Collection
    
        Open FileName For Random Access Read As InFile Len = Len(FR)
            Get #InFile, , Signature ' read file signature
            If Signature = FILE_SIGNATURE Then ' check the file signature
                ' read all records
                Do While Not EOF(InFile)
                    Get #InFile, , FR
                    
                    Set obj = New clsRECORDS
                    obj.LN = Trim$(FR.LN)
                    obj.FN = Trim$(FR.FN)
                    obj.MI = Trim$(FR.MI)
                    FileOpen.Add obj
                Loop
            
                FileOpen.Remove FileOpen.Count ' remove the last record
            Else
                ' if signature not found then display error message
                MsgBox "File format error!", vbCritical
            End If
        Close
    Else
        MsgBox "File not found!", vbInformation
    End If
End Function

'=============================================================================================
Private Sub FileSave(FileName As String, Data As Collection)
    Dim i As Integer
    Dim InFile As Integer
    Dim FR As FILE_RECORD
    On Error Resume Next ' turns off error handling
    
    If Dir$(FileName) Then Kill FileName ' if file found then delete
    
    InFile = FreeFile ' file number
    
    Open FileName For Random Access Write As InFile Len = Len(FR) ' open file for writing
        Put #InFile, , FILE_SIGNATURE ' write file signature
        For i = 1 To Data.Count ' number of records
            FR.LN = Data(i).LN
            FR.FN = Data(i).FN
            FR.MI = Data(i).MI
            Put #InFile, , FR
        Next i
    Close InFile ' close and saves the records
End Sub

