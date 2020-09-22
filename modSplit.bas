Attribute VB_Name = "modSplit"
'This function will split the file into many files
Function SplitFile(Filename As String, Filesize As Long) As Boolean
    On Error GoTo handelsplit
    'Some variables
    Dim lSizeOfFile As Long, iCountFiles As Integer
    Dim iNumberOfFiles As Integer, lSizeOfCurrentFile As Long
    Dim sBuffer As String '10Kb buffer
    Dim sRemainBuffer As String, lEndPart As Long
    Dim lSizeToSplit As Long, sHeader As String * 16
    Dim iFileCounter As Integer, sNewFilename As String
    Dim lWhereInFileCounter As Long
    
    Open Filename For Binary As #1
    'Get the lenght of the file (in bytes)
    lSizeOfFile = LOF(1)
    lSizeToSplit = Filesize * 1024
    
    'If the file is smaller then the variable "Filesize" then tell the user
    If lSizeOfFile <= lSizeToSplit Then
        Close #1
        SplitFile = False
        MsgBox "This file is smaller than the selected split size!", 16, "Error!"
        Exit Function
    End If
    
    'Check if file isn't alread split
    sHeader = Input(16, #1)
    Close #1

    'If the header is "SPLITIT" then is the file already splitted
    If Mid$(sHeader, 1, 7) = "SPLITIT" Then
        MsgBox "This file is alread split!"
        SplitFile = False
        Exit Function
    End If
    'Open the File
    Open Filename For Binary As #1
    'Get the length of the file
    lSizeOfFile = LOF(1)
    'calculate the size of the variable "lSizeToSplit"
    lSizeToSplit = Filesize * 1024
    'Reset the FileCounter
    iCountFiles = 0
    'calculate the number of files after splitting
    iNumberOfFiles = (lSizeOfFile \ lSizeToSplit) + 1
    'Set the sHeader data
    sHeader = "SPLITIT" & Format$(iFileCounter, "000") & Format$(iNumberOfFiles, "000") & Right$(Filename, 3)
    'Set the new filename
    sNewFilename = Left$(Filename, Len(Filename) - 3) & Format$(iFileCounter, "000")
    'create the new file
    Open sNewFilename For Binary As #2
    'now write the header
    Put #2, , sHeader
    'Now set the size of the current file
    lSizeOfCurrentFile = Len(sHeader)
    'While the end of file (EOF) is not reached...
    While Not EOF(1)
    '...create the Buffer
        sBuffer = Input(10240, #1)
        'Set the size of the current file
        lSizeOfCurrentFile = lSizeOfCurrentFile + Len(sBuffer)
        'If the size of the current file is greater then the lSizeToSplit variable
        If lSizeOfCurrentFile > lSizeToSplit Then
            'Write last bit
            lEndPart = Len(sBuffer) - (lSizeOfCurrentFile - lSizeToSplit) + Len(sHeader)
            Put #2, , Mid$(sBuffer, 1, lEndPart)
            'close the file
            Close #2
            'Make new file
            iFileCounter = iFileCounter + 1
            'Set the header
            sHeader = "SPLITIT" & Format$(iFileCounter, "000") & Format$(iNumberOfFiles, "000") & Right$(Filename, 3)
            'make new filename
            sNewFilename = Left$(Filename, Len(Filename) - 3) & Format$(iFileCounter, "000")
            'Open the file
            Open sNewFilenfame For Binary As #2
            'Write the header
            Put #2, , sHeader
            'Put Rest of buffer read
            Put #2, , Mid$(sBuffer, lEndPart + 1)
            'Set the new size
            lSizeOfCurrentFile = Len(sHeader) + (Len(sBuffer) - lEndPart)
        Else
            Put #2, , sBuffer
        End If
    Wend
    
    'close all files
    Close #2
    Close #1
    SplitFile = True
    Exit Function
handelsplit:
    SplitFile = False
    'On error display error
    MsgBox Err.Description, 16, "Error #" & Err.Number
    Exit Function
End Function


Function AssembleFile(Filename As String) As Boolean
    On Error GoTo handelassemble
    'some variables
    Dim sHeader As String * 16
    Dim sBuffer As String '10Kb buffer
    Dim sFileExt As String, iNumberOfFiles As Integer
    Dim iCurrentFileNumber As Integer
    Dim iCounter As Integer, sTempFilename As String
    Dim sNewFilename As String
    'Open the file
    Open Filename For Binary As #1
    'get the header
    sHeader = Input(Len(sHeader), #1)
    
    'Check if it's a splitted file
    If Mid$(sHeader, 1, 7) <> "SPLITIT" Then
        MsgBox "This is Not a split file!"
        AssembleFile = False
        'Leave the function
        Exit Function
    Else
        'The first file is a split file ok
        'Read the header values
        iCurrentFileNumber = Val(Mid$(sHeader, 8, 3))
        'Get number of splitted files
        iNumberOfFiles = Val(Mid$(sHeader, 11, 3))
        'Get extension
        sFileExt = Mid$(sHeader, 14, 3)
        'If the Filenumber is not 0 then tell the user
        If iCurrentFileNumber <> 0 Then
            MsgBox "This is Not the first file In the sequence!" & vbCrLf & "The extension of the file must be .000!"
            AssembleFile = False
            'Leave the function
            Exit Function
        End If
    End If
    'Close the file
    Close #1
    'Set the new filename
    sNewFilename = Left$(Filename, Len(Filename) - 3) & sFileExt
    'Set the CurrentFile variable in frmMain
    frmMain.CurrentFile = sNewFilename
    'Create the assembled file
    Open sNewFilename For Binary As #2
    'Assemble files
    For iCounter = 0 To iNumberOfFiles - 1
        'Create a Tempfile
        sTempFilename = Left$(Filename, Len(Filename) - 3) & Format$(iCounter, "000")
        'Open the Tempfile
        Open sTempFilename For Binary As #1
        'Get header
        sHeader = Input(Len(sHeader), #1)
        'If the header isn't "SPLITIT" Then tell the user
        If Mid$(sHeader, 1, 7) <> "SPLITIT" Then
            MsgBox "This is not a splitted file!" & sTempFilename
            AssembleFile = False
            'Leave the function
            Exit Function
        End If
        'Set the current filenumber
        iCurrentFileNumber = Val(Mid$(sHeader, 8, 3))
        'If the Filenumber isn't the same as iCounter the tell the user
        If iCurrentFileNumber <> iCounter Then
            MsgBox "The file '" & sTempFilename & "' is out of sequence!"
            AssembleFile = False
            'Close all files
            Close #2
            Close #1
            'Leave the function
            Exit Function
        End If
        'While End Of File is not reached...
        While Not EOF(1)
            'Create the buffer
            sBuffer = Input(10240, #1)
            'Put the bufferdata into the file
            Put #2, , sBuffer
        Wend
        'Close the file
        Close #1
        'Show the progressbar
        ShowProgress frmMain.picProgress, iCounter, 0, iNumberOfFiles - 1, False
        'Get next file
    Next iCounter
    'Close the file
    Close #2
    AssembleFile = True
    'Leave the function.
    'Else we would get an error
    Exit Function
handelassemble:
    AssembleFile = False
    'Display the errornumber
    MsgBox Err.Description, 16, "Error #" & Err.Number
    Exit Function
End Function

Function ShowProgress(picProgress As PictureBox, ByVal Value As Long, ByVal Min As Long, ByVal Max As Long, Optional ByVal bShowProzent As Boolean = True)
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
    ' check if AutoReadraw is True
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
End Function

