Attribute VB_Name = "modMainFunctions"
'Force variable declaration
Option Explicit

'Function, which will detect the main informations about the system
'NOTE: The "lpVersionInformation" argument is declared "As Any" in order to accept the "OSVERSIONINFOEX" type, too
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
'Function, which will detect a 64 Bit process
'NOTE: This API is available only under Windows XP and Windows .NET
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
'Function, which will detect the current process name
'We need this function together with the "IsWow64Process" API call
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Functions, which will open the Registry, get a value from it and close the Registry
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Function, which we need for the detection of Windows' language
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

'Main Type, which will hold the informations we need in order to detect the Windows version
'NOTE: This Type is supported by all 32 Bit Windows versions
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'Type, which will hold the extended informations we need in order to detect the Windows version
'NOTE: This Type is supported only by Windows 2000 and higher
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

'The constants we need for the informations, which we send and recieve from the Registry
Private Const REG_SZ As Long = 1
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const KEY_QUERY_VALUE As Long = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8&
Private Const KEY_NOTIFY As Long = &H10&
Private Const KEY_READ As Long = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const ERROR_SUCCESS As Long = 0
Private Const csKey As String = "SYSTEM\CurrentControlSet\Control\ProductOptions"
Private Const csName As String = "ProductType"

'Global variables, which we need in some functions
Private Informations As OSVERSIONINFO
Private ExtendedInformations As OSVERSIONINFOEX

'Detects the Windows "Build" number
Public Function GetBuildNumber() As Integer

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    GetBuildNumber = Informations.dwBuildNumber

End Function

'Detects extended informations about the Windows version:
'Windows 95/98: Only the first byte is used in order to identify the Windows version (A, B or C)
'Windows NT/2000/XP/.NET: The whole string is used to identify the Windows version (Service Pack 5...)
Public Function GetExtendedInformations() As String

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    GetExtendedInformations = Trim$(Informations.szCSDVersion)

End Function

'Detects the Windows "Major" number
Public Function GetMajorNumber() As Integer

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    GetMajorNumber = Informations.dwMajorVersion

End Function

'Detects the Windows "Minor" number
Public Function GetMinorNumber() As Integer

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    GetMinorNumber = Informations.dwMinorVersion

End Function

'Detects the Windows "Platform" number
'0 = Windows 32 Bit
'1 = Windows 95 based system
'2 = Windows NT based system
'3 = Windows CE
Public Function GetPlatformNumber() As Integer

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    GetPlatformNumber = Informations.dwPlatformId

End Function

'Detects the Windows "Platform" number
'Same as "GetPlatformNumber", but it returns the platform name, not number
Public Function GetPlatformType() As String

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    If Informations.dwPlatformId = 0 Then
        GetPlatformType = "32 Bit Windows"
      ElseIf Informations.dwPlatformId = 1 Then
        GetPlatformType = "Windows 95"
      ElseIf Informations.dwPlatformId = 2 Then
        GetPlatformType = "Windows NT"
      ElseIf Informations.dwPlatformId = 3 Then
        GetPlatformType = "Windows CE"
      Else
        GetPlatformType = "Unknown Platform Type"
    End If

End Function

'Detects the product family (Workstation, Server, Advanced Server...)
'NOTE: This function will work only on Windows NT based operating systems
Public Function GetProductFamily() As String

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    If Informations.dwPlatformId = 2 And Informations.dwMajorVersion > 4 Then
        ExtendedInformations.dwOSVersionInfoSize = Len(ExtendedInformations)
        GetVersionEx ExtendedInformations
        If ExtendedInformations.wProductType And 1 Then
            GetProductFamily = "Workstation/Professional/Home"
          ElseIf ExtendedInformations.wProductType And 3 Then
            If ExtendedInformations.wSuiteMask And 128 Then
                GetProductFamily = "DataCenter Server"
              ElseIf ExtendedInformations.wSuiteMask And 2 Then
                GetProductFamily = "Advanced/Enterprise Server"
              Else
                GetProductFamily = "Server"
            End If
          Else
            GetProductFamily = "Unknown Product Family"
        End If
      ElseIf Informations.dwPlatformId = 2 And Informations.dwMajorVersion < 5 Then
        'For Windows NT 4.0 and bellow, we have to "ask" the Registry
        If GetWindowsNTFamily = "WinNT" Then
            GetProductFamily = "Workstation"
          ElseIf GetWindowsNTFamily = "ServerNT" Then
            GetProductFamily = "Server"
          Else
            GetProductFamily = "Advanced Server"
        End If
      Else
        GetProductFamily = "Unknown Product Family"
    End If

End Function

'Detect the full Windows 2000 version ("Windows 2000 %PRODUCT_FAMILY% Service Pack %SERVICE_PACK_NUMBER%")
Public Function GetWindows2000Version() As String

  Dim Windows2000Informations As OSVERSIONINFO
  Dim ExtendedWindows2000Informations As OSVERSIONINFOEX

    Windows2000Informations.dwOSVersionInfoSize = Len(Windows2000Informations)
    ExtendedWindows2000Informations.dwOSVersionInfoSize = Len(ExtendedWindows2000Informations)
    GetVersionEx Windows2000Informations

    If Windows2000Informations.dwPlatformId = 2 And Windows2000Informations.dwMajorVersion = 5 And Windows2000Informations.dwMinorVersion = 0 And Windows2000Informations.dwBuildNumber = 2195 Then
        GetVersionEx ExtendedWindows2000Informations
        If ExtendedWindows2000Informations.wProductType = 1 Then
            GetWindows2000Version = "Windows 2000 Professional"
          ElseIf ExtendedWindows2000Informations.wProductType = 3 Then
            If ExtendedWindows2000Informations.wSuiteMask = 128 Then
                GetWindows2000Version = "Windows 2000 DataCenter Server"
              ElseIf ExtendedWindows2000Informations.wSuiteMask = 2 Then
                GetWindows2000Version = "Windows 2000 Advanced Server"
              Else
                GetWindows2000Version = "Windows 2000 Server"
            End If
        End If
        If ExtendedWindows2000Informations.wServicePackMajor = 1 Then
            GetWindows2000Version = GetWindows2000Version & " Service Pack 1"
          ElseIf ExtendedWindows2000Informations.wServicePackMajor = 2 Then
            GetWindows2000Version = GetWindows2000Version & " Service Pack 2"
          ElseIf ExtendedWindows2000Informations.wServicePackMajor = 3 Then
            GetWindows2000Version = GetWindows2000Version & " Service Pack 3"
          Else
            GetWindows2000Version = GetWindows2000Version & " " & Trim$(Windows2000Informations.szCSDVersion)
        End If
      Else
        GetWindows2000Version = "No Windows 2000 System"
    End If

End Function

'Detect the Windows 95 based (Windows 95, 98 and ME) version
Public Function GetWindows95BasedSystem() As String

  Dim Windows95BasedSystemInformations As OSVERSIONINFO

    Windows95BasedSystemInformations.dwOSVersionInfoSize = Len(Windows95BasedSystemInformations)
    GetVersionEx Windows95BasedSystemInformations

    If Windows95BasedSystemInformations.dwPlatformId = 1 And Windows95BasedSystemInformations.dwMinorVersion < 10 Then
        GetWindows95BasedSystem = GetWindows95Version
      ElseIf Windows95BasedSystemInformations.dwPlatformId = 1 And Windows95BasedSystemInformations.dwMinorVersion = 10 Then
        GetWindows95BasedSystem = GetWindows98Version
      ElseIf Windows95BasedSystemInformations.dwPlatformId = 1 And Windows95BasedSystemInformations.dwMinorVersion >= 90 Then
        GetWindows95BasedSystem = GetWindowsMEVersion
      Else
        GetWindows95BasedSystem = "No Windows 95 Based System"
    End If

End Function

'Detect the full Windows 95 version ("Windows 95 Service Pack %SERVICE_PACK_NUMBER%/%RELEASE_VERSION%")
Public Function GetWindows95Version() As String

  Dim Windows95Informations As OSVERSIONINFO

    Windows95Informations.dwOSVersionInfoSize = Len(Windows95Informations)
    GetVersionEx Windows95Informations

    If Windows95Informations.dwPlatformId = 1 And Windows95Informations.dwMinorVersion < 10 Then
        If Windows95Informations.dwBuildNumber = 950 Then
            GetWindows95Version = "Windows 95"
          ElseIf Windows95Informations.dwBuildNumber > 950 And Windows95Informations.dwBuildNumber <= 1080 Then
            GetWindows95Version = "Windows 95 Service Pack 1"
          ElseIf Windows95Informations.dwBuildNumber >= 1111 And Windows95Informations.dwBuildNumber <= 1211 Then
            GetWindows95Version = "Windows 95 OSR-2"
          ElseIf Windows95Informations.dwBuildNumber >= 1212 And Windows95Informations.dwBuildNumber <= 1213 Then
            GetWindows95Version = "Windows 95 OSR-2.1"
          ElseIf Windows95Informations.dwBuildNumber > 1213 Then
            GetWindows95Version = "Windows 95 OSR-2.5"
          Else
            GetWindows95Version = "Unknown Windows 95 Version"
        End If
      Else
        GetWindows95Version = "No Windows 95 System"
    End If

End Function

'Detect the full Windows 98 version ("Windows 98 Service Pack %SERVICE_PACK_NUMBER%/%RELEASE_VERSION%")
Public Function GetWindows98Version() As String

  Dim Windows98Informations As OSVERSIONINFO

    Windows98Informations.dwOSVersionInfoSize = Len(Windows98Informations)
    GetVersionEx Windows98Informations

    If Windows98Informations.dwPlatformId = 1 And Windows98Informations.dwMinorVersion = 10 Then
        If Windows98Informations.dwBuildNumber = 1998 Then
            GetWindows98Version = "Windows 98"
          ElseIf Windows98Informations.dwBuildNumber >= 1999 And Windows98Informations.dwBuildNumber <= 2182 Then
            GetWindows98Version = "Windows 98 Service Pack 1"
          ElseIf Windows98Informations.dwBuildNumber >= 2183 Then
            GetWindows98Version = "Windows 98 SE"
          Else
            GetWindows98Version = "Unknown Windows 98 Version"
        End If
      Else
        GetWindows98Version = "No Windows 98 System"
    End If

End Function

'Detect the Windows CE version
Public Function GetWindowsCEVersion() As String

  Dim WindowsCEInformations As OSVERSIONINFO

    WindowsCEInformations.dwOSVersionInfoSize = Len(WindowsCEInformations)
    GetVersionEx WindowsCEInformations

    If WindowsCEInformations.dwPlatformId = 3 Then
        If WindowsCEInformations.dwMajorVersion = 1 Then
            GetWindowsCEVersion = "Windows CE 1.0"
          ElseIf WindowsCEInformations.dwMajorVersion = 2 Then
            If WindowsCEInformations.dwMinorVersion = 0 Then
                GetWindowsCEVersion = "Windows CE 2.0"
              ElseIf WindowsCEInformations.dwMinorVersion = 1 Then
                GetWindowsCEVersion = "Windows CE 2.1"
              Else
                GetWindowsCEVersion = "Unknown Windows CE 2 Version"
            End If
          ElseIf WindowsCEInformations.dwMajorVersion = 3 Then
            GetWindowsCEVersion = "Windows CE 3.0"
          ElseIf WindowsCEInformations.dwMajorVersion = 4 Then
            GetWindowsCEVersion = "Windows CE 4.0"
          ElseIf WindowsCEInformations.dwMajorVersion = 5 Then
            GetWindowsCEVersion = "Windows CE .NET"
          Else
            GetWindowsCEVersion = "Unknown Windows CE Version"
        End If
      Else
        GetWindowsCEVersion = "No Windows CE System"
    End If

End Function

'Detect the Windows Code Name string
'NOTE1: This function isn't included in the Windows version and cannot be retrieved by the Windows API
'NOTE2: I have not included code names for Windows CE, as every Windows CE version (for PocketPC, for mobile phones...) has its own code name
Public Function GetWindowsCodeName() As String

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    If Informations.dwPlatformId = 0 Then
        GetWindowsCodeName = "Unknown Code Name"
      ElseIf Informations.dwPlatformId = 1 Then
        If Informations.dwMinorVersion < 10 Then
            GetWindowsCodeName = "Chicago"
          ElseIf Informations.dwMinorVersion = 10 Then
            GetWindowsCodeName = "Memphis"
          ElseIf Informations.dwMinorVersion >= 90 Then
            GetWindowsCodeName = "Millenium"
        End If
      ElseIf Informations.dwPlatformId = 2 Then
        If Informations.dwMajorVersion = 3 Then
            GetWindowsCodeName = "Daytona"
          ElseIf Informations.dwMajorVersion = 4 Then
            GetWindowsCodeName = "Cairo"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 0 Then
            GetWindowsCodeName = "2K"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 1 Then
            GetWindowsCodeName = "Whistler"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 2 Then
            GetWindowsCodeName = ".NET"
          Else
            GetWindowsCodeName = "Unknown Code Name"
        End If
      Else
        GetWindowsCodeName = "Unknown Code Name"
    End If

End Function

'Detect the Windows folder, based on the PATH (Environment)
Public Function GetWindowsFolder() As String

    GetWindowsFolder = Environ$("WINDIR")

End Function

'Determine the Windows language
Public Function GetWindowsLanguage(UseEnglishName As Boolean) As String

  Dim Buffer As String
  Dim Ret As String

    Buffer = String$(256, 0)
    If UseEnglishName Then
        Ret = GetLocaleInfo(ByVal GetSystemDefaultLCID, 4097, Buffer, Len(Buffer))
      Else
        Ret = GetLocaleInfo(ByVal GetSystemDefaultLCID, 4, Buffer, Len(Buffer))
    End If
    If Ret > 0 Then
        GetWindowsLanguage = Left$(Buffer, Ret - 1)
      Else
        GetWindowsLanguage = ""
    End If

End Function

'Detect the Windows ME version
Public Function GetWindowsMEVersion() As String

  Dim WindowsMEInformations As OSVERSIONINFO

    WindowsMEInformations.dwOSVersionInfoSize = Len(WindowsMEInformations)
    GetVersionEx WindowsMEInformations

    If WindowsMEInformations.dwPlatformId = 1 And WindowsMEInformations.dwMinorVersion >= 90 Then
        If WindowsMEInformations.dwBuildNumber > 2493 Then
            GetWindowsMEVersion = "Windows ME"
          Else
            GetWindowsMEVersion = "Unknown Windows ME Version"
        End If
      Else
        GetWindowsMEVersion = "No Windows ME System"
    End If

End Function

'Detect the Windows name
Public Function GetWindowsName() As String

    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    If Informations.dwPlatformId = 0 Then
        GetWindowsName = "32 Bit Windows"
      ElseIf Informations.dwPlatformId = 1 Then
        If Informations.dwMinorVersion < 10 Then
            GetWindowsName = "Windows 95"
          ElseIf Informations.dwMinorVersion = 10 Then
            GetWindowsName = "Windows 98"
          ElseIf Informations.dwMinorVersion >= 90 Then
            GetWindowsName = "Windows ME"
          Else
            GetWindowsName = "Unknown Windows Name"
        End If
      ElseIf Informations.dwPlatformId = 2 Then
        If Informations.dwMajorVersion < 5 Then
            GetWindowsName = "Windows NT"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 0 Then
            GetWindowsName = "Windows 2000"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 1 Then
            GetWindowsName = "Windows XP"
          ElseIf Informations.dwMajorVersion = 5 And Informations.dwMinorVersion = 2 Then
            GetWindowsName = "Windows .NET"
          Else
            GetWindowsName = "Unknown Windows Name"
        End If
      ElseIf Informations.dwPlatformId = 3 Then
        GetWindowsName = "Windows CE"
      Else
        GetWindowsName = "Unknown Windows Name"
    End If

End Function

'Detect the full Windows .NET version ("Windows .NET %PRODUCT_FAMILY% Service Pack %SERVICE_PACK_NUMBER%")
Public Function GetWindowsNETVersion() As String

  Dim WindowsNETInformations As OSVERSIONINFO
  Dim ExtendedWindowsNETInformations As OSVERSIONINFOEX

    WindowsNETInformations.dwOSVersionInfoSize = Len(WindowsNETInformations)
    ExtendedWindowsNETInformations.dwOSVersionInfoSize = Len(ExtendedWindowsNETInformations)
    GetVersionEx WindowsNETInformations

    If WindowsNETInformations.dwPlatformId = 2 And WindowsNETInformations.dwMajorVersion = 5 And WindowsNETInformations.dwMinorVersion = 2 Then
        GetVersionEx ExtendedWindowsNETInformations
        If ExtendedWindowsNETInformations.wSuiteMask = 2 Then
            GetWindowsNETVersion = "Windows .NET Enterprise Server"
          Else
            GetWindowsNETVersion = "Windows .NET Server"
        End If
        If ExtendedWindowsNETInformations.wServicePackMajor = 1 Then
            GetWindowsNETVersion = GetWindowsNETVersion & " Service Pack 1"
          ElseIf ExtendedWindowsNETInformations.wServicePackMajor = 2 Then
            GetWindowsNETVersion = GetWindowsNETVersion & " Service Pack 2"
          Else
            GetWindowsNETVersion = GetWindowsNETVersion & " " & Trim$(WindowsNETInformations.szCSDVersion)
        End If
      Else
        GetWindowsNETVersion = "No Windows .NET System"
    End If

End Function

'Detect the Windows NT based (Windows NT, 2000, XP and .NET) version
Public Function GetWindowsNTBasedSystem() As String

  Dim WindowsNTBasedSystemInformations As OSVERSIONINFO

    WindowsNTBasedSystemInformations.dwOSVersionInfoSize = Len(WindowsNTBasedSystemInformations)
    GetVersionEx WindowsNTBasedSystemInformations

    If WindowsNTBasedSystemInformations.dwPlatformId = 2 And WindowsNTBasedSystemInformations.dwMajorVersion <= 4 Then
        GetWindowsNTBasedSystem = GetWindowsNTVersion
      ElseIf WindowsNTBasedSystemInformations.dwPlatformId = 2 And WindowsNTBasedSystemInformations.dwMajorVersion = 5 And WindowsNTBasedSystemInformations.dwMinorVersion = 0 Then
        GetWindowsNTBasedSystem = GetWindows2000Version
      ElseIf WindowsNTBasedSystemInformations.dwPlatformId = 2 And WindowsNTBasedSystemInformations.dwMajorVersion = 5 And WindowsNTBasedSystemInformations.dwMinorVersion = 1 Then
        GetWindowsNTBasedSystem = GetWindowsXPVersion
      ElseIf WindowsNTBasedSystemInformations.dwPlatformId = 2 And WindowsNTBasedSystemInformations.dwMajorVersion = 5 And WindowsNTBasedSystemInformations.dwMinorVersion = 2 Then
        GetWindowsNTBasedSystem = GetWindowsNETVersion
      Else
        GetWindowsNTBasedSystem = "No Windows NT Based System"
    End If

End Function

'Detect the Windows NT Product Family based on the informations we get from the Windows Registry
Public Function GetWindowsNTFamily() As String

  Dim lHKey As Long
  Dim lRtn As Long
  Dim lpcbData As Long
  Dim sReturnedString As String

    lRtn = RegOpenKeyEx(HKEY_LOCAL_MACHINE, csKey, 0&, KEY_READ, lHKey)
    If lRtn = ERROR_SUCCESS Then
        lpcbData = 1024
        sReturnedString = Space$(lpcbData)
        lRtn = RegQueryValueEx(lHKey, csName, ByVal 0&, REG_SZ, sReturnedString, lpcbData)
        If lRtn = ERROR_SUCCESS Then
            GetWindowsNTFamily = Left$(sReturnedString, lpcbData - 1)
        End If
        RegCloseKey lHKey
    End If

End Function

'Detect the full Windows NT version ("Windows NT %PRODUCT_FAMILY% Service Pack %SERVICE_PACK_NUMBER%")
Public Function GetWindowsNTVersion() As String

  Dim WindowsNTInformations As OSVERSIONINFO

    WindowsNTInformations.dwOSVersionInfoSize = Len(WindowsNTInformations)
    GetVersionEx WindowsNTInformations

    If WindowsNTInformations.dwPlatformId = 2 And WindowsNTInformations.dwMajorVersion > 2 And WindowsNTInformations.dwMajorVersion < 5 Then
        If WindowsNTInformations.dwMajorVersion = 3 Then
            If WindowsNTInformations.dwMinorVersion = 0 Then
                GetWindowsNTVersion = "Windows NT 3.0"
              ElseIf WindowsNTInformations.dwMinorVersion = 1 Then
                GetWindowsNTVersion = "Windows NT 3.1"
              ElseIf WindowsNTInformations.dwMinorVersion = 5 Then
                GetWindowsNTVersion = "Windows NT 3.5"
              ElseIf WindowsNTInformations.dwMinorVersion = 51 Then
                GetWindowsNTVersion = "Windows NT 3.51"
              Else
                GetWindowsNTVersion = "Unknown Windows NT 3 Version"
            End If
          ElseIf WindowsNTInformations.dwMajorVersion = 4 Then
            GetWindowsNTVersion = "Windows NT 4.0"
        End If
        If GetWindowsNTFamily = "WinNT" Then
            GetWindowsNTVersion = GetWindowsNTVersion & " Workstation"
          ElseIf GetWindowsNTFamily = "ServerNT" Then
            GetWindowsNTVersion = GetWindowsNTVersion & " Server"
          Else
            GetWindowsNTVersion = GetWindowsNTVersion & " Advanced Server"
        End If
        If Not Trim$(WindowsNTInformations.szCSDVersion) = "" Then
            If Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "1" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 1"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "2" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 2"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "3" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 3"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "4" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 4"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "5" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 5"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "6" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 6/6a"
              ElseIf Trim$(Mid$(WindowsNTInformations.szCSDVersion, 14, 1)) = "7" Then
                GetWindowsNTVersion = GetWindowsNTVersion & " Service Pack 7"
              Else
                GetWindowsNTVersion = "Windows NT 4.0 " & Trim$(WindowsNTInformations.szCSDVersion)
            End If
        End If
      Else
        GetWindowsNTVersion = "No Windows NT System"
    End If

End Function

'Detect the full Windows version
Public Function GetWindowsVersion() As String

  Dim WindowsInformations As OSVERSIONINFO

    WindowsInformations.dwOSVersionInfoSize = Len(WindowsInformations)
    GetVersionEx WindowsInformations

    If WindowsInformations.dwPlatformId = 0 Then
        GetWindowsVersion = "Unknown 32 Bit Windows Version"
      ElseIf WindowsInformations.dwPlatformId = 1 Then
        GetWindowsVersion = GetWindows95BasedSystem
      ElseIf WindowsInformations.dwPlatformId = 2 Then
        GetWindowsVersion = GetWindowsNTBasedSystem
      ElseIf WindowsInformations.dwPlatformId = 3 Then
        GetWindowsVersion = GetWindowsCEVersion
      Else
        GetWindowsVersion = "Unable To Get System Informations"
    End If

End Function

'Detect the full Windows XP version ("Windows XP %PRODUCT_FAMILY% Service Pack %SERVICE_PACK_NUMBER%")
Public Function GetWindowsXPVersion() As String

  Dim WindowsXPInformations As OSVERSIONINFO
  Dim ExtendedWindowsXPInformations As OSVERSIONINFOEX

    WindowsXPInformations.dwOSVersionInfoSize = Len(WindowsXPInformations)
    ExtendedWindowsXPInformations.dwOSVersionInfoSize = Len(ExtendedWindowsXPInformations)
    GetVersionEx WindowsXPInformations

    If WindowsXPInformations.dwPlatformId = 2 And WindowsXPInformations.dwMajorVersion = 5 And WindowsXPInformations.dwMinorVersion = 1 And WindowsXPInformations.dwBuildNumber = 2600 Then
        GetVersionEx ExtendedWindowsXPInformations
        If ExtendedWindowsXPInformations.wSuiteMask And 512 Then
            GetWindowsXPVersion = "Windows XP Home"
          Else
            GetWindowsXPVersion = "Windows XP Professional"
        End If
        If ExtendedWindowsXPInformations.wServicePackMajor = 1 Then
            GetWindowsXPVersion = GetWindowsXPVersion & " Service Pack 1"
          Else
            GetWindowsXPVersion = GetWindowsXPVersion & " " & Trim$(WindowsXPInformations.szCSDVersion)
        End If
      Else
        GetWindowsXPVersion = "No Windows XP System"
    End If

End Function

'Check if the OS is a 64 Bit system
'NOTE: Only Windows XP and .NET can be 64 Bit OSs
Public Function Is64BitSystem() As Boolean

  Dim Wow64Process As Long

    On Error GoTo ErrorHandler

    IsWow64Process GetCurrentProcess, Wow64Process
    Is64BitSystem = Wow64Process <> 0

Exit Function

ErrorHandler:

End Function

'Determine the Service Pack version based on the extended API call
'NOTE: Works only under Windows 2000 and above
Public Function NewGetServicePackNumber() As String

    ExtendedInformations.dwOSVersionInfoSize = Len(ExtendedInformations)
    GetVersionEx ExtendedInformations
    NewGetServicePackNumber = ExtendedInformations.wServicePackMajor & "." & ExtendedInformations.wServicePackMinor

End Function

'Determine the Service Pack version based on the standard API call
'NOTE: Works on all Windows NT based systems
Public Function OldGetServicePackNumber() As String

    Err.Raise 6
    Informations.dwOSVersionInfoSize = Len(Informations)
    GetVersionEx Informations
    OldGetServicePackNumber = Trim$(Mid$(Informations.szCSDVersion, 14, 1))

End Function


'####################################################################################

'I have used If..Then..Else instead of Select Case because there are some problems with the ExtendedInformations.wProductType and ExtendedInformations.wSuiteMask
'You CANNOT use "Case 512... Blah Blah Blah" for example, you have to use "Case Is = 512... Blah Blah Blah"
'I don't know exactly why

'####################################################################################

'This is a very basic example on how to get the Windows version without the API:
'Public Function GetWindowsVersion() As String
'
'    If Environ$("OS") = "" Then
'        GetWindowsVersion = "Windows 95/98/ME"
'      Else
'        If Environ$("PROGRAMFILES") = "" Then
'            GetWindowsVersion = "Windows NT 3/4)"
'          Else
'            GetWindowsVersion = "Windows 2000/XP/.NET"
'        End If
'    End If
'
'End Function

'####################################################################################

