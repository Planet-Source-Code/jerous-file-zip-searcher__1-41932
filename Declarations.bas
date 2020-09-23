Attribute VB_Name = "Declarations"
Option Explicit

Public Const AppName = "File Searcher"
'gfx
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


'Drive recognition
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOVABLE = 2
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000


'Folder selection
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type


'Properties of a file
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public fso As New FileSystemObject 'needs a reference to Microsoft Scripting Runtime

Public Const SEE_MASK_INVOKEIDLIST As Long = &HC
Public Const SEE_MASK_FLAG_NO_UI As Long = &H400

Public Const SW_SHOWNORMAL As Long = 1

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'read in registry

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Enum HKeyTypes
    hkey_classes_root = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
    Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number

'extract icon from a file
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'file deletion
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_SHIFT = &H10
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FOF_ALLOWUNDO = &H40

Function ExtractPicture(file As String) As String
Dim Descr$, Ext$, PicPath$, IconIndex%
Dim Icon As Long, pic As Picture
Ext = "." + Right(file, Len(file) - InStrRev(file, "."))
Descr = GetString(hkey_classes_root, Ext, "")
'retrieve path for defaulticon
PicPath = GetString(hkey_classes_root, Descr + "\defaulticon", "")

If Descr = "" Then Exit Function
If PicPath = "" Then
'  'try to get the icon from the default open program
'  PicPath = GetString(hkey_classes_root, Descr + "\shell\open\command", "")
'  If PicPath = "" Then Exit Function
'  Do
'    PicPath = Mid(PicPath, 1, Len(PicPath) - 1)
'  Loop Until dir(PicPath) <> "" Or Len(PicPath) <= 3
'  PicPath = Trim(PicPath)
'  If PicPath = "" Then
'    Exit Function 'no icon found, so exit function
'  End If
'  Extract PicPath, 0
'  ExtractPicture = PicPath
Else
  If InStr(PicPath, ",") = 0 Then Exit Function
  'get number of the icon in the file
  IconIndex = Right(PicPath, Len(PicPath) - InStrRev(PicPath, ","))
  'sometimes i have here a iconindex smaller than 0
  If IconIndex < 0 Then IconIndex = 0
  'extract the path
  PicPath = Mid(PicPath, 1, InStrRev(PicPath, ",") - 1)
  Extract PicPath, IconIndex
  'return the path (got from the register) to the function
  ExtractPicture = PicPath + "," + Trim(Str(IconIndex))
End If
End Function
Private Function Extract(Path$, IconIndex)
Dim Icon As Long
'clear picture
Form1.pic = LoadPicture("")
'extract the icon, handle stored in Icon
Icon = ExtractIcon(Form1.pic.hdc, Path, IconIndex)
'draw the icon on the picturebox
DrawIcon Form1.pic.hdc, 0, 0, Icon
'refresh picture
Form1.pic.Refresh
Form1.pic.Picture = Form1.pic.Image
'free resources used by the icon
'it may be destroyed because it's already stored in the picture
DestroyIcon Icon
End Function
Function GetFileType(file As String) As String
Dim Descr$, Ext$
Ext = "." + Right(file, Len(file) - InStrRev(file, "."))
Descr = GetString(hkey_classes_root, Ext, "")
GetFileType = GetString(hkey_classes_root, Descr, "")
If GetFileType = "" Then
  GetFileType = UCase(Mid(Ext, 2)) + " file"
End If
End Function
Public Function GetString(hKey As HKeyTypes, strPath As String, strValue As String)
'read a string from the registry
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim lValueType As Long

RegOpenKey hKey, strPath, keyhand
RegQueryValueEx keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function

