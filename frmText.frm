VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmText 
   Caption         =   "Text-View"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox Text 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmText.frx":0000
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Make the textbox flat
    MakeFlat Text.hWnd
    'Make the form flat
    MakeFlat Me.hWnd
    'Resize the Textbox
    Text.Width = Me.ScaleWidth
    Text.Height = Me.ScaleHeight
    'Open the Text
    Text.LoadFile App.Path + "\temp\" + frmProp.txtName.Text
    Me.Caption = Me.Caption + "      " + frmProp.txtName.Text
End Sub

Private Sub Form_Resize()
    'Resize the Textbox
    Text.Width = Me.ScaleWidth
    Text.Height = Me.ScaleHeight
End Sub

Sub OpenText(ByVal cText As TextBox, ByVal Filename As String)
    'A variable
    Dim strFileData As String
    'Open the file
    Open Filename For Binary As #1
        'Get the FileData
        strFileData = String(LOF(1), 0)
        'Get the file
        Get #1, 1, strFileData
    'Close the file
    Close #1
    'Put the Text into the textbox
    cText.Text = strFileData
    'Clear the variable (saves memory)
    strFileData = ""
End Sub

