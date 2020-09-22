VERSION 5.00
Begin VB.Form frmProp 
   BackColor       =   &H00CD651F&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdText 
      Caption         =   "Open as Text"
      Height          =   400
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00733E0D&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00733E0D&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Image pic 
      Height          =   1695
      Left            =   30
      Stretch         =   -1  'True
      Top             =   40
      Width           =   1575
   End
   Begin VB.Label LblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
      Height          =   195
      Left            =   1815
      TabIndex        =   4
      Top             =   480
      Width           =   600
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
      Height          =   195
      Left            =   1695
      TabIndex        =   3
      Top             =   120
      Width           =   720
   End
   Begin VB.Label NoPic 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Picture or no supported picture"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Result As String
'This API will convert bytes into KB or MB
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String


Private Sub cmdExtract_Click()
    'Get Outputfolder
    Result = BrowseForFolder("Select Outputfolder:")
    'The file is already extracted, just copy it
    frmMain.FSO.CopyFile App.Path + "\temp\" + txtName.Text, Result, True
    'Finished!
    MsgBox "Finished!", vbInformation, "Finished!"
End Sub

Private Sub cmdText_Click()
    'Open it as a textfile
    frmText.Show (vbModal), frmMain
End Sub

Private Sub Form_Load()
    'First make all controls flat
    MakeFlat OK.hWnd
    MakeFlat cmdExtract.hWnd
    MakeFlat txtName.hWnd
    MakeFlat txtSize.hWnd
    MakeFlat Me.hWnd
    MakeFlat cmdText.hWnd
    'Get the original Filename
    Result = frmMain.Files.list(frmMain.Files.ListIndex)
    txtName.Text = Result
    'If it's a picture, display it!
    If Right(Result, 3) = "bmp" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "gif" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "jpg" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "jpeg" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "emf" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "wmf" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "dib" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "ico" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    ElseIf Right(Result, 3) = "cur" Then
        pic.Picture = LoadPicture(App.Path + "\temp\" + txtName.Text)
    End If
    'If it isn't a picture, the Imagecontrol will be unvisible and the label
    'in the back will be displayed!
    
    'Get the lenght of the file in bytes
    txtSize.Text = FormatKB(FileLen(App.Path + "\temp\" + txtName.Text))
End Sub

Private Sub OK_Click()
    'Delete the extracted file
    Kill App.Path + "\temp\" + txtName.Text
    'Unload the Form
    Unload Me
    'If the Form is minimized, restore it
    frmMain.Show
End Sub

'This will convert bytes into KB or MB
Public Function FormatKB(ByVal Amount As Long) As String
    'Some variables
    Dim Buffer As String
    Dim Result As String
    'Create the buffer
    Buffer = Space$(255)
    'Format the ByteSize
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    'convert the bytes into KB or MB
    If InStr(Result, vbNullChar) > 1 Then FormatKB = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function

