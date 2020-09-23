Attribute VB_Name = "Ext"
'Fileaccess for save/load

Option Explicit
Declare Function OSGetPrivateProfileInt Lib "KERNEL32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function OSGetPrivateProfileSection Lib "KERNEL32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function OSGetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function OSWritePrivateProfileSection Lib "KERNEL32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function OSWritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function OSGetProfileInt Lib "KERNEL32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Declare Function OSGetProfileSection Lib "KERNEL32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function OSGetProfileString Lib "KERNEL32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Declare Function OSWriteProfileSection Lib "KERNEL32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Declare Function OSWriteProfileString Lib "KERNEL32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Declare Function ExitWindowsEx Lib "user32" _
    (ByVal uFlags As Long, ByVal dwReserved As Long)
    
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
Public Const HTLEFT = 10
Public Const HTRIGHT = 11

Public Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   Dim szTmp                     As String
   Dim nRet                      As Long
   If (IsNull(szEntry)) Then
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)
End Function

Public Sub WritePrivateProfile(ByVal szSection As String, ByVal szEntry As Variant, ByVal vValue As Variant, ByVal szFileName As String)
   Dim nRet                      As Long
   If (IsNull(szEntry)) Then
      nRet = OSWritePrivateProfileString(szSection, 0&, 0&, szFileName)
   ElseIf (IsNull(vValue)) Then
      nRet = OSWritePrivateProfileString(szSection, CStr(szEntry), 0&, szFileName)
   Else
      nRet = OSWritePrivateProfileString(szSection, CStr(szEntry), CStr(vValue), szFileName)
   End If
End Sub
Public Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   Dim szTmp                    As String
   Dim nRet                     As Long
   If (IsNull(szEntry)) Then
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)
End Function
Public Function FileReal(Filename As String) As Boolean
If Filename = "" Then
FileReal = False
Exit Function
End If

If LCase(Dir(Filename)) = LCase(GetFileName(Filename)) Then
FileReal = True
Else
FileReal = False
End If
End Function
Function GetFileName(fname As String) As String
    Dim i As Long

    On Error Resume Next
    For i = Len(fname) To 1 Step -1
        If Mid(fname, i, 1) = "\" Then
            Exit For
        End If
    Next i
    GetFileName = Trim(Mid(fname, i + 1))
End Function

Public Function CheckBMP(Filename As String) As Boolean
If Right(Filename, 3) = "bmp" Then
CheckBMP = True
Exit Function
End If
If Right(Filename, 3) = "jpg" Then
CheckBMP = True
Exit Function
End If
If Right(Filename, 3) = "gif" Then
CheckBMP = True
Exit Function
End If
If Right(Filename, 4) = "jpeg" Then
CheckBMP = True
Exit Function
End If
CheckBMP = False
End Function
    



Public Sub MoveForm(f As Form)
    ReleaseCapture
    SendMessage f.hWnd, WM_NCLBUTTONDOWN, 2, 0
End Sub

