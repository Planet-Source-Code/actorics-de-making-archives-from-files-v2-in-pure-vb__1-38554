VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00733E0D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ArchiveMaker V2"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Align           =   2  'Align Bottom
      ForeColor       =   &H00CF760A&
      Height          =   225
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   8610
      TabIndex        =   6
      Top             =   6600
      Width           =   8670
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   8160
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Bar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00BE7D56&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   8640
      TabIndex        =   1
      Top             =   0
      Width           =   8670
      Begin VB.CommandButton cmdSplit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Split Archive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   8
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5784
         TabIndex        =   7
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdExtract 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Extract"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4368
         TabIndex        =   5
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdMake 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2952
         TabIndex        =   4
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1536
         TabIndex        =   3
         Top             =   50
         Width           =   1335
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   50
         Width           =   1335
      End
   End
   Begin VB.ListBox Files 
      Height          =   5130
      ItemData        =   "frmMain.frx":0442
      Left            =   0
      List            =   "frmMain.frx":048B
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   720
      Width           =   8655
   End
   Begin VB.Menu mnuProp 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "Show properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi!
'This is my third upload on PSC.
'And it's fucking good code!!!

'The new searchfunction isn't written by me.
'I found it on PSC.
'So thanks to the man who wrote this.
'It's a very good function.

'The code stores files in one file, so it can be extracted later.
'For example it can be used to extract bitmaps at runtime so the user can't
'steal them.
'The Listbox supports Drag&Drop, so you can easyily drag some files into the Listbox
'and store them in a archive.
'YOU WILL GET AN ERROR IF YOU TRY TO PUT A FILE WITH A HIGHER SIZE THEN
'200 MB FILE INTO AN ARCHIVE!!!!!!!
'FOR BIG FILES USE THE OLD VERSION (also included in an archive in the App.Path
'named "old.dat")!
'If the archive is damaged, you can extract it although,
'because if you extract a damaged archive,
'the damaged area of the file will be added to the last extracted file.
'You can also open an existing archive and look what's in it and then
'extract it!!!!!
'In every second line is a comment.
'So the code is easy to understand.

'A man wondered why some archives are compressed.
'I think, it's because of Microsoft's FileSystem (Clusters, 32KB).
'If you know why, please mail me at:
                            '   irnchen@web.de
'If you got questions please mail me at:
                            '   irnchen@web.de
                            
'P.S.: ____....----====PLEASE VOTE FOR THIS GOOD CODE!!!!!!====----....____
      '(Or I will not make any updates!!!!!!!!!!!)

'If you want to use the code in your app, only copy all functions and subs in the
'(General) section in your app

'How to use them will show you the following code:

Public CurrentFile As String
Dim Result As String
Dim Result2 As Integer
Dim Result3 As Long
Public FSO As New FileSystemObject
'Create a FileSystemObject for copy files

Private Sub cmdExtract_Click()
    'Error-Handling
    If Not CurrentFile = vbNullString Then
        Result = modFolder.BrowseForFolder("extract to:")
        'Extract the Files
        Result2 = ExtractFilesFromArchive(CurrentFile, Result)
        'Clear the progressbar
        picProgress.Cls
        'Tell the user that the extraction is finished
        MsgBox "Finished!!!", vbInformation, "Finished!!!"
    Else
        'If no archive is open, tell the user
        MsgBox "No archive open!!", vbCritical, "Error!"
        Exit Sub
    End If
End Sub

Private Sub cmdMake_Click()
'Some variables
Dim i As Integer
Dim nLenFileName As Integer
Dim File As String
Dim num As Integer
Dim p As String
Dim pass As String
    'Check if the Listbox is empty
    If Files.ListCount = 0 Then
        MsgBox "You can not make an archive from an empty list!", vbCritical, "Error"
        Exit Sub
    End If
    'Show the SaveDialog
    Dialog.DialogTitle = "Save a archive"
    Dialog.Filter = "Archives (*.dat)|*.dat|All Files (*.*)|*.*"
    Dialog.ShowSave
    'Set CurrentFile
    CurrentFile = Dialog.Filename
    'MAKE THE ARCHIVE!!
    MakeArchiveFromList Files '(-=NEW!!!!=-)
    'Tell the user that the archive is finished!
    MsgBox "Finished!!!", vbInformation, "Finished!!!"
End Sub

Private Sub cmdNew_Click()
    'Clear all Items in the Listbox
    Files.Clear
    CurrentFile = ""
End Sub

Private Sub cmdOpen_Click()
'Error Handling
  On Error Resume Next
'Some variables
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim fileData As String
  Dim File As String
  Dim nFiles As Long
  Dim i As Long
  Dim sDestDir As String

    'Show the OpenDialog for to open a archive
    Dialog.DialogTitle = "Open a archive"
    Dialog.Filter = "Archives (*.dat)|*.dat|Split Files (*.000)|*.000|All Files (*.*)|*.*"
    Dialog.ShowOpen
    'Check If Filename is empty
If Not Dialog.Filename = vbNullString Then
    'Set CurrentFile
    CurrentFile = Dialog.Filename
    'Check if it is a splitted archive
    If Right(CurrentFile, 3) = "000" Then
        'If it is a splitted archive, assemble it.
        If AssembleFile(CurrentFile) Then
        End If
    End If
    'The assemble function sets automaticilly the variable CurrentFile
    'So we can open the Archive
    ReadFilesIntoList Files
  'Let's try to open it!!
Else
    'If Filename is empty then exit the sub
    Exit Sub
End If
End Sub

Private Sub cmdSearch_Click()
    Result = InputBox("File to search for:", "Search:")
    SearchInList Result, Files
End Sub

Private Sub cmdSplit_Click()
    'This will split the current archive in some files
    'First ask the user, if he want to extract a splitted archive
    If MsgBox("Do you want to extract a splitted archive?", vbYesNo, "extract?") = vbYes Then
        'If yes, then assemble the file into one file
        Dialog.DialogTitle = "select splitted file"
        Dialog.Filter = "*.000|*.000"
        Dialog.ShowOpen
        'Put all splitted files into one file
        If AssembleFile(Dialog.Filename) Then
        End If
        'Now ask for a output dir
        Result = BrowseForFolder("Select an output folder:")
        'Extract the archive
        Result3 = ExtractFilesFromArchive(CurrentFile, Result)
        'Tell the user that it's finished
        MsgBox "Finished!", vbInformation, "Finished!"
    Else
    'Check if a File is open
    If Not CurrentFile = "" Then
    'Ask the user, what sizes he want
    Result3 = InputBox("In which size do you want to split the file?", "Size")
        If SplitFile(CurrentFile, Result3) Then
            'If it is finished, then tell the user
            MsgBox "Finished!"
    Else
        'If no File is open, tell the user
        MsgBox "No File selected!", vbCritical, "Error"
    End If
    End If
    End If
End Sub

Private Sub Files_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next 'Starts then error handle
    'Removes a ListItem
    If KeyCode = vbKeyDelete Then
        Files.RemoveItem Files.ListIndex
    End If
End Sub

Private Sub Files_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show popupmenu on Rightclick
    If Button = 2 Then
        'Show the popupmenu
        PopupMenu mnuProp
    End If
End Sub

Private Sub Files_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'If the user drags files into the Listbox, we'll add them
    Dim i As Integer
    'Get number of Files
    For i = 1 To Data.Files.count
        'Now add them to the Listbox
        Files.AddItem Data.Files(i)
    Next i
End Sub

Private Sub Form_Load()
    'Resize the Listbox
    Files.Width = Me.ScaleWidth
    Files.Height = Me.ScaleHeight - Bar.ScaleHeight - picProgress.ScaleHeight
    'Make the Buttons flat
    MakeFlat cmdNew.hWnd
    MakeFlat cmdOpen.hWnd
    MakeFlat cmdMake.hWnd
    MakeFlat cmdExtract.hWnd
    MakeFlat cmdSearch.hWnd
    MakeFlat cmdSplit.hWnd
    MakeFlat picProgress.hWnd
    'Show the Form
    Me.Show
End Sub

'This Function will save the Files to an archive
Public Function SaveFilesToArchive(ByVal sPath As String, ByVal sArchiv As String, Optional ByVal sPattern As String = "*.*") As Long
'Error Handling:
'Some variables
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim fileData As String
  Dim File() As String
  Dim nFiles As Long
  Dim i As Long
  Dim lngUBound As Long

  ' Add backslash to the path
  If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"

  ' Get all files in the directory
  nFiles = 0

  DirName = Dir(sPath & sPattern, vbNormal)
  While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
      nFiles = nFiles + 1
      'Get files
      If nFiles > lngUBound Then lngUBound = 2 * nFiles
      ReDim Preserve File(lngUBound)
      File(nFiles) = DirName
    End If
    DirName = Dir
  Wend
  ReDim Preserve File(nFiles)

  ' If archiv exists already, delete it
  If Dir(sArchiv) <> "" Then Kill sArchiv

  ' Now save all files to the archive
  F = FreeFile
  Open sArchiv For Binary As #F

  ' Set number of files
  Put #F, , nFiles

  For i = 1 To nFiles
    ' Save filename
    nLenFileName = Len(File(i))
    Put #F, , nLenFileName
    Put #F, , File(i)

    ' Read filedata
    n = FreeFile
    Open sPath + File(i) For Binary As #n
    fileData = Space$(LOF(n))
    Get #n, , fileData
    Close #n

    ' Save filedata to the archive
    nLenFileData = Len(fileData)
    Put #F, , nLenFileData
    Put #F, , fileData
    ' Progress
    ShowProgress frmMain.picProgress, i, 1, nFiles
    DoEvents
  Next i
  Close #F
  
  SaveFilesToArchiv = nFiles
  
  Exit Function
  'Error Handling:
Error:
    MsgBox "Error while saving." & vbCrLf & "Check if you got enough memory and if the directory isn't locked.", vbCritical, "Error!"
End Function

' extract all files to the outputfolder
Public Function ExtractFilesFromArchive(ByVal sArchiv As String, ByVal sDestDir As String) As Long
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim fileData As String
  Dim File As String
  Dim nFiles As Long
  Dim i As Long
  
  ' check if Archiv exists
  If Dir(sArchiv) = "" Then
    MsgBox "The archive does not exist!", 16
    Exit Function
  End If
  
  ' add backslash to the path
  If Right$(sDestDir, 1) <> "\" Then _
    sDestDir = sDestDir + "\"
  
  ' Open the archive
  F = FreeFile
  Open sArchiv For Binary As #F
  
  ' Get number of Icons in the archive
  Get #F, , nFiles
  
  For i = 1 To nFiles
    ' get original filenames
    Get #F, , nLenFileName
    File = Space$(nLenFileName)
    Get #F, , File
    
    ' Read filedata
    Get #F, , nLenFileData
    fileData = Space$(nLenFileData)
    Get #F, , fileData
    
    ' Save file in "DestDir"
    n = FreeFile
    Open sDestDir + File For Output As #n
    Print #n, fileData;
    Close #n
    
    ' Progress
    ShowProgress frmMain.picProgress, i, 1, nFiles
    DoEvents
  Next i
  'Close the file
  Close #F
  
  ExtractFilesFromArchiv = nFiles
  Exit Function
End Function
' Progressbar
Private Sub ShowProgress(picProgress As PictureBox, ByVal Value As Long, ByVal Min As Long, ByVal Max As Long, Optional ByVal bShowProzent As Boolean = True)
'Some variables
  Dim pWidth As Long
  Dim intProz As Integer
  Dim strProz As String
  
  ' colors
  Const progBackColor = &HCF760A
  Const progForeColor = &HCD651F
  Const progForeColorHighlight = vbWhite
  
  ' set Values
  If Value < Min Then Value = Min
  If Value > Max Then Value = Max
  
  ' calculate the percent
  If Max > 0 Then
    intProz = Int(Value / Max * 100 + 0.5)
  Else
    intProz = 100
  End If
    
  With picProgress
    ' check if AutoReadraw=True
    If .AutoRedraw = False Then .AutoRedraw = True
    
    ' clear the picturebox
    picProgress.Cls
    
    If Value > 0 Then
    
      ' calculate barwidth
      pWidth = .ScaleWidth / 100 * intProz
      
      ' Show bar
      picProgress.Line (0, 0)-(pWidth, .ScaleHeight), _
        progBackColor, BF
        
      ' show percent
      If bShowProzent Then
        strProz = CStr(intProz) & " %"
        .CurrentX = (.ScaleWidth - .TextWidth(strProz)) / 2
        .CurrentY = (.ScaleHeight - .TextHeight(strProz)) / 2
      
        ' If the width of the bar is higher then the font then use an other
        ' color for the font
        If pWidth >= .CurrentX Then
          .ForeColor = progForeColorHighlight
        Else
        ' If not then let the fontcolor as it is
          .ForeColor = progForeColor
        End If
        'Show the percent
        picProgress.Print strProz
      End If
    End If
  End With
End Sub

'This function will get the original FileTitle
Public Function GetFileTitle(ByVal sFilename As String) As String
Dim lPos As Long
    'Returns the position of the last occurrence of one string within another
    lPos = InStrRev(sFilename, "\")
    If lPos > 0 Then
        'If lPos is < then the number of chars in sFilename
        If lPos < Len(sFilename) Then
            'Then trim the Path from the FileTitle
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            'If not then set the Function = ""
            GetFileTitle = ""
        End If
      Else
        GetFileTitle = sFilename
    End If
End Function

Public Function ReadFilesIntoList(ByVal list As ListBox)
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim fileData As String
  Dim File As String
  Dim nFiles As Long
  Dim i As Long
  
 'Clear the Listbox
  list.Clear
 'Tell F that it is a emtpy file
  F = FreeFile
  'Open the Current File
  Open CurrentFile For Binary As #F
  ' Get number of Icons in the archive
  Get #F, , nFiles
  For i = 1 To nFiles
    ' get original filenames
    Get #F, , nLenFileName
    File = Space$(nLenFileName)
    Get #F, , File
    
    ' Read filedata
    Get #F, , nLenFileData
    fileData = Space$(nLenFileData)
    Get #F, , fileData
    
    ' Add File to the Listbox
    n = FreeFile
    list.AddItem File
    Next i
    Close #F
End Function

Public Function MakeArchiveFromList(ByVal list As ListBox)
'Error Handling
On Error GoTo Error
        'Copy all Files to a tempdir
        For i = 0 To list.ListCount - 1
        'Get the FileTitle
        p = GetFileTitle(list.list(i))
        'Now copy the file to the tempdir
        'If you want to store BIG files in a archive, it will take a while...
        FileCopy list.list(i), App.Path + "\temp\" & p
        ShowProgress picProgress, i, 1, list.ListCount - 1
        Next i
        'Now save all Files from the tempdir to the archive!!!
        Result = SaveFilesToArchive(App.Path + "\temp", CurrentFile, "*.*")
        'Delete the tempdir
        FSO.DeleteFolder App.Path + "\temp"
        'Create the new empty tempdir
        MkDir App.Path + "\temp"
        'Clear the progressbar
        picProgress.Cls
        'Exit the function, else we'll get an error
        Exit Function
Error:
        'Check if the opened file is an archive
        If Right(CurrentFile, 3) = "dat" Then
            MsgBox "You opened a archive. You can't make an archive from an archive!!", vbCritical, "Error"
            Exit Function
        Else
            MsgBox "Unknown Error", vbCritical, "Error"
            Exit Function
        End If
End Function

'Search-Function
Private Function SearchInList(SearchText As String, list As ListBox)
'Some variables
    Dim count As Integer
    Dim sSource As String
    Dim tmpSource As String
    Dim sTarget As String
    Dim iLast As Integer
    Dim iCurrent As Integer
    Dim done As Boolean
    'first get the searchstring
    sSource = LCase(SearchText)
    'iLast is 0
    iLast = 0
    For count = 0 To (list.ListCount - 1)
    ' Get all Listboxindexes
        sTarget = LCase(list.list(count))
        ' Get the item
        done = False
        ' Set Done = False
            iCurrent = 0 ' Set iCurrent = 0
            tmpSource = sSource ' Set tmpSource = sSource
            Do
            'Go in a Do-Loop
                If Left(sTarget, 1) = Left(tmpSource, 1) Then
                    ' Current first char's matched
                    iCurrent = iCurrent + 1
                    ' Increase number of matching char's
                    sTarget = Right(sTarget, Len(sTarget) - 1)
                    ' Trim the text down by one
                    tmpSource = Right(tmpSource, Len(tmpSource) - 1)
                    ' Trim the text down by one
                Else
                    done = True
                    ' No more matching char's
                    End If
                Loop While done = False
                ' Check result:
                If iLast > iCurrent Then
                    ' Last item had more matching char's, se
                    '     lect it
                    list.ListIndex = count - 1
                    ' Set new list index
                    count = list.ListCount - 1
                    ' Exit for/next Loop
                ElseIf iLast = iCurrent And iLast > 0 Then
                    ' Current has just as many matching char
                    's as last, select last pos.
                    list.ListIndex = count - 1
                    ' Set new list index
                    count = list.ListCount - 1
                    ' Exit for/next Loop
                ElseIf iCurrent > iLast And count = (list.ListCount - 1) Then
                    ' Last line in list and line with most m
                    '     atching char's, select it
                    list.ListIndex = count
                    ' Set new list index
                Else
                    iLast = iCurrent
                    ' Update iLast
                End If
            Next count
            ' Loop search
End Function

'This is a FUCKING GOOD FUNCTION, YOU CAN EXTRACT ONLY ONE FILE FROM THE ARCHIVE!
Public Function ExtractOneFileFromArchive(ByVal sArchive As String, ByVal Filename As String, ByVal sDestDir As String) As String
On Error GoTo Error:
  'Some variables
  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim fileData As String
  Dim File As String
  Dim nFiles As Long
  Dim i As Long
  
  ' check if Archiv exists
  If Dir(sArchive) = "" Then
    MsgBox "The archive does not exist!", 16
    Exit Function
  End If
  
  ' add backslash to the path
  If Right$(sDestDir, 1) <> "\" Then _
    sDestDir = sDestDir + "\"
  
  ' Open the archive
  F = FreeFile
  Open sArchive For Binary As #F
  
  ' Get number of Icons in the archive
  Get #F, , nFiles
  
  For i = 1 To nFiles
    ' get original filenames
    Get #F, , nLenFileName
    File = Space$(nLenFileName)
    Get #F, , File
    
    ' Read filedata
    Get #F, , nLenFileData
    fileData = Space$(nLenFileData)
    Get #F, , fileData
    
    'If the file is the same as "Filename", extract it
    If File = Filename Then
        ' Save file in "DestDir"
        n = FreeFile
        'Create the file
        Open sDestDir + File For Output As #n
            'Write in the data
            Print #n, fileData;
        'Close the file
        Close #n
    End If
    ShowProgress frmMain.picProgress, i, 1, nFiles
    DoEvents
  Next i
  'Close the file
  Close #F
  'Leave the function
  Exit Function
Error:
    MsgBox "This is not a valid file!!!", vbCritical, "ERROR!"
    ExtractOneFileFromArchive = "error"
    Exit Function
End Function

Private Sub mnuExtract_Click()
    'Get the outputfolder
    Result = BrowseForFolder("Select an outputfolder:")
    'Extract the selected File from the archive (-=NEW!!!!=-)
    ExtractOneFileFromArchive CurrentFile, Files.list(Files.ListIndex), Result
End Sub

Private Sub mnuProperty_Click()
    'If no file is selected then make an error
    If Files.list(Files.ListIndex) = "" Then
        MsgBox "No file selected!", vbCritical, "Error"
    Else
        'If we get an error while extracting we'll exit the sub
        If ExtractOneFileFromArchive(CurrentFile, Files.list(Files.ListIndex), App.Path + "\temp") = "error" Then
            Exit Sub
        Else
            'If we get no error, show the Propertyform (-=NEW!!!!=-)
            frmProp.Show (vbModal), Me
        End If
    End If
End Sub

'If you didn't read the top of the source, then do it by NOW!!!!

'So now you know how to use it.
'Have fun with it.
'Totally FREE source.
'BUT VOTE!!!!!!!!!!!!!!!!!!

'(c) by Irnchen

