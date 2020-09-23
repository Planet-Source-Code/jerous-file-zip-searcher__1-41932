Attribute VB_Name = "ZipDll"
'==============================================================================
'Richsoft Computing 2001
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'
'For latest information about this and other projects please visit my website:
'www.richsoftcomputing.btinternet.co.uk
'
'If you would like to make any comments/suggestions then please e-mail them to
'richsoftcomputing@btinternet.co.uk
'==============================================================================

'Declarations of the library functions
'Public Declare Function AddFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Public Declare Function ExtractFile Lib "zipit.dl" (ByVal ZipFilename As String, ByVal Filename As String, ByVal ExtrDir As String, ByVal UseDirInfo As Boolean, ByVal Overwrite As Boolean, ByVal Action As Integer) As Boolean
'Public Declare Function DeleteFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal Filename As String) As Boolean


'Class that holds the file info
Public Version As Integer
Public Flag As Integer
Public CompressionMethod As Integer
Public ZipFileDateTime As String
Public CRC32 As Long
Public CompressedSize As Long
Public UncompressedSize As Long
Public FileNameLength As Integer
Public ExtraFieldLength As Integer
'Public Filename As String

'Zip file format type
Type ZipFile
  Version As Integer                    ': WORD;
  Flag As Integer                       ': WORD;
  CompressionMethod As Integer          ': WORD;
  Time As Integer                       ': WORD;
  Date As Integer                       ': WORD;
  CRC32 As Long                      ': Longint;
  CompressedSize As Long             ': Longint;
  UncompressedSize As Long           ': Longint;
  FileNameLength As Integer             ': WORD;
  ExtraFieldLength As Integer           ': WORD;
  Filename As String                 ': String;
End Type

'Zip file constants
Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50

'App constants
Public Const APP_TITLE = "Richsoft Zipit 1.0"

'File dates/times functions and types
Public Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FileTime) As Long
Public Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


'==============================================================================
'Richsoft Computing 2001
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'
'For latest information about this and other projects please visit my website:
'www.richsoftcomputing.btinternet.co.uk
'
'If you would like to make any comments/suggestions then please e-mail them to
'richsoftcomputing@btinternet.co.uk
'==============================================================================

'Set up the private atrributes
Private ZipFilename As String
'Private CompLevel As ZipLevel
Private DOS83Format As Boolean
Private Recurse As Boolean

'Set up the file collection
Public Archive As Collection

'Actions
Public Enum ZipAction
    zipDefault = 1
    zipFreshen = 2
    zipUpdate = 3
End Enum

Public Property Get Filename() As String
    Filename = ZipFilename
End Property

Public Property Let Filename(ByVal New_Filename As String)
    Dim R As Long
    Dim i As Long
    'Called when the filename is updated
    ZipFilename = New_Filename
    'Read in the contents of the file
    R = Read
End Property

Public Function GetEntry(ByVal Index As Long) As ZipFileEntry
    Set GetEntry = Archive(Index)
End Function
Public Function GetEntryNum() As Long
    GetEntryNum = Archive.Count
End Function


Public Function ParsePath(Path As String) As String
    'Takes a full file specification and returns the path
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            'Add the correct path separator for the input
            If Mid$(Path, A, 1) = "\" Then
                ParsePath = LCase$(Left$(Path, A - 1) & "\")
            Else
                ParsePath = LCase$(Left$(Path, A - 1) & "/")
            End If
            Exit Function
        End If
    Next A
End Function

Public Function ParseFilename(ByVal Path As String) As String
    'Takes a full file specification and returns the path
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            ParseFilename = Mid$(Path, A + 1)
            Exit Function
        End If
    Next A
    ParseFilename = Path
End Function

Public Function Read() As Long
    'Reads the archive and places each file into a collection
    Dim Sig As Long
    Dim ZipStream As Integer
    Dim Res As Long
    Dim zFile As ZipFile
    Dim Name As String
    Dim i As Integer
    Dim Counter%
    'If the filename is empty return a empty file list
    If ZipFilename = "" Then
        Read = 0
        'Remove any files still in the list
        For i = Archive.Count To 1 Step -1
            Archive.Remove i
        Next i
        Exit Function
    End If
    
    'Clears the collection
    'begin
'    Set Archive = New Collection
'    Archive.Clear
    For i = Archive.Count To 1 Step -1
        Archive.Remove i
    Next i
    
    'Opens the archive for binary access
    ZipStream = FreeFile
    Open ZipFilename For Binary As ZipStream
    'Loop through archive
    Do While True
        Get ZipStream, , Sig
        'See if the file header has been found
              If Sig = LocalFileHeaderSig Then
                    'Read each part of the file header
                    Get ZipStream, , zFile.Version
                    Get ZipStream, , zFile.Flag
                    Get ZipStream, , zFile.CompressionMethod
                    Get ZipStream, , zFile.Time
                    Get ZipStream, , zFile.Date
                    Get ZipStream, , zFile.CRC32
                    Get ZipStream, , zFile.CompressedSize
                    Get ZipStream, , zFile.UncompressedSize
                    Get ZipStream, , zFile.FileNameLength
                    Get ZipStream, , zFile.ExtraFieldLength
                    'Get the filename
                    'Set up a empty string so the right number of
                    'bytes is read
                    Name = String$(zFile.FileNameLength, " ")
                    Get ZipStream, , Name
                    zFile.Filename = Mid$(Name, 1, zFile.FileNameLength)
                    'Move on through the archive
                    'Skipping extra space, and compressed data
                    Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                    'Add the fileinfo to the collection
                    AddEntry zFile
              Else
              'Debug.Print Sig
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    'All the filenames have been found so
                    'exit the loop
                    Exit Do
                'End
                Else
                If Sig = EndCentralDirSig Then
                    'Exit the loop
                    Exit Do
                End If
                End If
            End If
            Counter = Counter + 1
            If Counter > 200 Then
              DoEvents
              Counter = 0
              If Form1.StopClicked Or Form1.Unloading Then
                Close
                Exit Function
              End If
            End If
            
        Loop
        'Close the archive
        Close ZipStream
        'Return the number of files in the archive
        Read = Archive.Count

    'Fire the update event
'    RaiseEvent OnArchiveUpdate
End Function

Private Sub AddEntry(zFile As ZipFile)
    Dim xFile As New ZipFileEntry
    'Adds a file from the archive into the collection
    '**It does not add entry that are just folders**
    If ParseFilename(zFile.Filename) <> "" Then
        xFile.Version = zFile.Version
        xFile.Flag = zFile.Flag
        xFile.CompressionMethod = zFile.CompressionMethod
        xFile.CRC32 = zFile.CRC32
        xFile.FileDateTime = GetDateTime(zFile.Date, zFile.Time)
        xFile.CompressedSize = zFile.CompressedSize
        xFile.UncompressedSize = zFile.UncompressedSize
        xFile.FileNameLength = zFile.FileNameLength
        xFile.Filename = zFile.Filename
        xFile.ExtraFieldLength = zFile.ExtraFieldLength
    End If
    Archive.Add xFile
End Sub

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string

    Dim R As Long
    Dim FTime As FileTime
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String

    'Convert the dos stamp into a file time
    R = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Convert the file time into a standard time
    R = FileTimeToSystemTime(FTime, Sys)

    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond

    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function
