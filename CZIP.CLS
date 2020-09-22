VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260

'Zip file constants
Private Const LocalFileHeaderSig = &H4034B50
Private Const CentralFileHeaderSig = &H2014B50
Private Const EndCentralDirSig = &H6054B50

'Actions
Public Enum ZipAction
    zipDefault = 1
    zipFreshen = 2
    zipUpdate = 3
End Enum

'Compression Level
Public Enum ZipLevel
    zipStore = 0
    zipLevel1 = 1
    zipSuperFast = 2
    zipFast = 3
    zipLevel4 = 4
    zipNormal = 5
    zipLevel6 = 6
    zipLevel7 = 7
    zipLevel8 = 8
    zipMax = 9
End Enum
'----------------------------------------------------------------
'Zip file format type
Private Type ZipFile
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

Private Type ZipFileEntry
    Version As Integer
    Flag As Integer
    CompressionMethod As Integer
    FileDateTime As String
    CRC32 As Long
    CompressedSize As Long
    UncompressedSize As Long
    FileNameLength As Integer
    ExtraFieldLength As Integer
    Filename As String
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'------------------------------------------------------------

'==============================================================================
'Declarations of the library functions
Private Declare Function AddFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Private Declare Function ExtractFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal ExtrDir As String, ByVal UseDirInfo As Boolean, ByVal Overwrite As Boolean, ByVal Action As Integer) As Boolean
Private Declare Function DeleteFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String) As Boolean
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Set up the private atrributes
Private CompLevel As ZipLevel
Private DOS83Format As Boolean
Private Recurse As Boolean

'Set up the file collection
Private Archive() As ZipFileEntry

'Events
Public Event OnZipProgress(ByVal Percentage As Integer, ByVal Filename As String)
Public Event OnZipComplete(ByVal Successful As Boolean)
Public Event OnUnzipProgress(ByVal Percentage As Integer, ByVal Filename As String)
Public Event OnUnzipComplete(ByVal Successful As Boolean)
Public Event OnDeleteProgress(ByVal Percentage As Integer, ByVal Filename As String)
Public Event OnDeleteComplete(ByVal Successful As Boolean)

'File dates/times functions and types
Private Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FileTime) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long

Private Function Add(ByVal ZipFileName As String, ByRef files As Variant, ByVal Action As ZipAction, ByVal StorePathInfo As Boolean, ByVal RecurseSubFolders As Boolean, ByVal UseDOS83 As Boolean, ByVal CompressionLevel As ZipLevel) As Long

    'Adds the specified files to the archive
    Dim ArchiveFilename                 As String
    Dim i                               As Long
    Dim Result                          As Long
    Dim FilesToAdd                      As Collection
    
    ArchiveFilename = ZipFileName

    'Check to see if there are any files in the archive
    'if not delete the file so there are not error messages

    If GetEntryNum = 0 Then
        If Dir$(ArchiveFilename) <> "" Then
            Kill ArchiveFilename
        End If
    End If

    'Find all the files to add, recursing folders if selected
    Set FilesToAdd = FindFiles(files)

    'Loop through the files adding them to the archive
    For i = 1 To FilesToAdd.Count
        Debug.Print "Trying to Add " & FilesToAdd(i)
        RaiseEvent OnZipProgress((100 * (i / (FilesToAdd.Count))), ParseFilename(FilesToAdd(i)))
        DoEvents
        If AddFile(ArchiveFilename, FilesToAdd(i), StorePathInfo, UseDOS83, Action, CompressionLevel) Then
            'File successfully extracted
            Result = Result + 1
        Else
            'File did not extract for some reason
        End If
    Next i

    RaiseEvent OnZipComplete(Result)
    'If any file was added update the archive
    If Result > 0 Then
        Read ZipFileName
    End If
    
End Function

Private Function FindFiles(files As Variant) As Variant

    'Return all the files (not directories) from the files array
    
    Dim res         As Collection
    Dim i           As Long
    
    Set res = New Collection

    For i = 1 To UBound(files, 2)
        If Not (files(1, i) = "D") Then
            res.Add files(0, i)
        End If
    Next i

    Set FindFiles = res

End Function

Public Function ExtractZipFile(ByVal ZipFileName As String, ByVal DestinationPath As String) As Long

    Dim files                   As Collection
    Dim i                       As Long
    Dim TotalSize               As Long
    
    Read ZipFileName 'first read the archive
    Set files = New Collection
    
    For i = 1 To GetEntryNum
        files.Add ParseFilename(GetFilename(i))
    Next i
    
    ExtractEX ZipFileName, files, zipDefault, False, True, DestinationPath
    Set files = Nothing
    
    RaiseEvent OnUnzipComplete(True)
    
End Function


Private Function ExtractEX(ByVal ZipFileName As String, ByRef files As Collection, ByVal Action As ZipAction, ByVal UsePathInfo As Boolean, ByVal Overwrite As Boolean, ByVal path As String) As Long

    'Extracts the files specified in the collection
    'Returns the number of files extracted
    
    Dim FilesToExtract              As Collection
    Dim ArchiveFilename             As String
    Dim i                           As Long
    Dim Result                      As Long
    
    ArchiveFilename = ZipFileName

    'First find the files which match the patterns
    'specified in the collection
    Set FilesToExtract = SelectFiles(files)

    'Extract each file in turn
    For i = 1 To FilesToExtract.Count
        Debug.Print "Trying to Extract " & FilesToExtract(i) & " to " & path
        RaiseEvent OnUnzipProgress((100 * (i / (FilesToExtract.Count))), ParseFilename(FilesToExtract(i)))
        DoEvents
        If ExtractFile(ArchiveFilename, CStr(FilesToExtract(i)), path, UsePathInfo, Overwrite, Action) Then
            'File successfully extracted
            Result = Result + 1
        Else
            'File did not extract for some reason
        End If
    Next i
    
    ExtractEX = Result

End Function

Private Function GetEntry(ByVal Index As Long) As ZipFileEntry
    GetEntry = Archive(Index)
End Function

Private Function GetFilename(ByVal Index As Long) As String
    GetFilename = Archive(Index).Filename
End Function

Private Function GetEntryNum() As Long
    GetEntryNum = UBound(Archive)
End Function

Private Function SelectFiles(files As Collection) As Collection
    'Selects files from a wildcard specification
    'Wildcards only corrispond to the filename and not the path
    Dim i               As Long
    Dim j               As Long
    Dim Result          As Collection
    
    
    Set Result = New Collection
    
    'Loop through the collection looking at each entry
    For i = 1 To files.Count
        'Loop through the files in the archive checking the pattern
        For j = 1 To GetEntryNum()
            'Check the pattern, ignoring case
            If LCase$(ParseFilename(GetEntry(j).Filename)) Like LCase$(files(i)) Then
                'Its a match so add it to the new collection
                Result.Add GetEntry(j).Filename
            End If
        Next j
    Next i
    Set SelectFiles = Result
    
End Function

Private Function ParseFilename(ByVal path As String) As String

    'Takes a full file specification and returns the path
    Dim i           As Long
    
    For i = Len(path) To 1 Step -1
        If Mid$(path, i, 1) = "\" Or Mid$(path, i, 1) = "/" Then
            ParseFilename = Mid$(path, i + 1)
            Exit Function
        End If
    Next i
    ParseFilename = path
    
End Function

Private Function Read(ByVal ZipFileName As String) As Long

    'Reads the archive and places each file into a collection
    
    Dim Sig                 As Long
    Dim ZipStream           As Integer
    Dim res                 As Long
    Dim zFile               As ZipFile
    Dim Name                As String
    Dim i                   As Integer

    'If the filename is empty return a empty file list
    If ZipFileName = "" Then
        Read = 0
        'Remove any files still in the list
        ReDim Archive(0)
        Exit Function
    End If

    'Archive.Clear;
    ReDim Archive(0)
    
    'Opens the archive for binary access
    ZipStream = FreeFile
    Open ZipFileName For Binary As #ZipStream
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
             If Sig = CentralFileHeaderSig Or Sig = 0 Then
                 'All the filenames have been found so
                 'exit the loop
                 Exit Do
             Else
                 If Sig = EndCentralDirSig Then
                     'Exit the loop
                     Exit Do
                 End If
             End If
         End If
    Loop
    
    'Close the archive
    Close #ZipStream
    'Return the number of files in the archive
    Read = UBound(Archive)
    
End Function

Private Sub AddEntry(zFile As ZipFile)

    Dim xFile       As ZipFileEntry
    Dim i           As Long
    
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
    i = UBound(Archive) + 1
    ReDim Preserve Archive(i)
    Archive(i) = xFile
    
End Sub

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string

    Dim r                   As Long
    Dim FTime               As FileTime
    Dim Sys                 As SYSTEMTIME
    Dim ZipDateStr          As String
    Dim ZipTimeStr          As String

    'Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Convert the file time into a standard time
    r = FileTimeToSystemTime(FTime, Sys)

    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond

    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function

Public Function MakeZipFile(ByVal SourcePath As String, ByVal ZipFileName As String) As Boolean

    Dim files           As Variant
    Dim i               As Long


    ReDim files(1, 0)
    GetAllFiles SourcePath, files, 1 'get all files recursively from folder

    Add ZipFileName, files, zipDefault, True, False, False, zipNormal

End Function

Private Function GetAllFiles(ByVal path As String, files As Variant, ByVal Index As Long) As Variant

    Dim i                   As Long
    Dim iCount              As Long

    GetAllFilesEX path, files

    iCount = UBound(files, 2)
    For i = Index To iCount
        If files(1, i) = "D" Then 'case its a directory
            GetAllFiles files(0, i), files, iCount + 1
        End If
    Next i

End Function


Private Sub GetAllFilesEX(ByVal path As String, files As Variant)

    Dim sFile           As String
    Dim res             As Boolean
    Dim i               As Long

    i = UBound(files, 2)

    sFile = Dir(path & "\*.*", vbDirectory Or vbNormal Or vbHidden Or vbArchive)
    Do While sFile <> ""
        Debug.Print sFile
        If sFile <> "." And sFile <> ".." Then
            i = i + 1
            ReDim Preserve files(1, i)
            files(0, i) = path & "\" & sFile
            If (GetAttr(path & "\" & sFile) And vbDirectory) = vbDirectory Then
                files(1, i) = "D" 'mark as directory
            End If
        End If
        sFile = Dir
    Loop

End Sub

Private Sub Class_Initialize()
    ReDim Archive(0)
End Sub
