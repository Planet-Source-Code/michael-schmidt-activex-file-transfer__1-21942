Attribute VB_Name = "Transfer"
Option Explicit
'========================================================================
'   Class:          VBTransfer          | (C) 2000 Michael A. Schmidt   |
'   Author:         Mike Schmidt        =================================
'   Date:           March 26, 2001
'   E-mail:         mikes@mtdmarketing.com
'========================================================================
'   References:     MSWINSCK.OCX    Microsoft Winsock
'                   SCRRUN.DLL      Microsoft Scripting Runtime (FSO)
'========================================================================
'   This class is used to send and receive files. The user may create
'   objects to receive and send files. Events include such as transfer
'   rate, finished, error, cancellation.
'========================================================================
'   Bugs: If multiple files are sent/received to the same person, the
'   control will only do one transfer at a time, aka it will stop the
'   previous transfer until the latest one finishes. Not sure why it
'   does this, if you have any nfo, please mail me!
'========================================================================
' Data Seperator
Public Const NewCom As String = "§"


'====================================
'   GrabFilename
'====================================
Public Function GrabFilename(FullPath As String)

    ' Pulls the filename from a full path and filename.
    ' Returns filename.
    GrabFilename = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))

End Function


'====================================
'   isValidFile
'====================================
Public Function isValidFile(ByVal iFile As String)
Dim FSO As New FileSystemObject

    ' Check to see if file exists.
    ' Return Boolean.
    isValidFile = FSO.FileExists(iFile)

End Function


'====================================
'   GetFileSize
'====================================
' Note This function will error if the
' remote FSO version (SCRIPTING) is old.
Public Function GetFileSize(ByVal iFile As String)
Dim FSO As New FileSystemObject
Dim FSOfile As File

    ' Get Size of File.
    ' Return File Size.
    Set FSOfile = FSO.GetFile(iFile)
    GetFileSize = FSOfile.Size

End Function


'====================================
'   GetOpenFileSize
'====================================
Public Function GetOpenFileSize(ByVal iFile As String)
Dim FileNumber

    ' - Open files return a different filesize
    ' - than just reading off the disk.
    
    ' - First grab an open filehandle (random = freefile)
    ' - Then open the file, grab the size (LOF) and
    ' - close the file.
    FileNumber = FreeFile
    Open iFile For Binary Access Read As #FileNumber
    GetOpenFileSize = LOF(FileNumber)
    Close #FileNumber


End Function


'====================================
'   FormatBytes
'====================================
Public Function FormatBytes(iBytes As Long) As String

    If iBytes < 1024 Then
        FormatBytes = iBytes & " b"
    ElseIf iBytes < 1048576 Then
        FormatBytes = Format(iBytes / 1024, "0.0") & " kb"
    Else 'If iBits < 1000000000 Then
        FormatBytes = Format(iBytes / 1048576, "0.00") & " mb"
    End If

End Function


'====================================
'   Word Function
'====================================
Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
' This function is used to parse data. Data is send as
' multiple commands in one packet. Each command is seperated
' by a special character. We retrieve specific commands by
' calling 'word' and specifying what seperates each 'word'.
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim X       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

'sSource = CSpace(sSource)

'find the nth word
X = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function

