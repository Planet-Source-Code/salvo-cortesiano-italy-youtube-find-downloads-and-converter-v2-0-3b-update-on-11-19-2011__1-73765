Attribute VB_Name = "modMain"

    Option Explicit
    
    Public DEFAULTVIDEOPATH As String
    
    ' .... Extend Combo Box
    Public Const CB_SETDROPPEDWIDTH = &H160
    
    ' .... Printers
    Private Const HWND_BROADCAST = &HFFFF&
    Private Const WM_WININICHANGE = &H1A
    Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
    Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

    
    Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGSTRUC) As Long

    Private Type PRINTDLGSTRUC
        lStructSize As Long: hWnd As Long: hDevMode As Long
        hDevNames As Long: hDC As Long
        flags As Long: nFromPage As Integer
        nToPage As Integer: nMinPage As Integer
        nMaxPage As Integer: nCopies As Integer
        hInstance As Long: lCustData As Long
        lpfnPrintHook As Long: lpfnSetupHook As Long
        lpPrintTemplateName As String: lpSetupTemplateName As String
        hPrintTemplate As Long: hSetupTemplate As Long
    End Type

    Private Const PD_ENABLEPRINTHOOK = &H1000
    Private Const PD_ENABLEPRINTTEMPLATE = &H4000
    Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    Private Const PD_ENABLESETUPHOOK = &H2000
    Private Const PD_ENABLESETUPTEMPLATE = &H8000
    Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    Private Const PD_NONETWORKBUTTON = &H200000
    Private Const PD_PRINTSETUP = &H40
    Private Const PD_USEDEVMODECOPIES = &H40000
    Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    Private Const PD_NOWARNING = &H80
    
    ' .... From DOS 8 to Win 32
    Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
    Const MAX_PATH = 260
    
    Public sStringText As String
    
    Public PID As Long
    
    Public strTemp As String
    
    Public estensione As String

    Public URLFlashVideo As String
    Public URLFlashIDVideo As String
    Public HOSTURL As String
    Public URLDownload As String
    Public URLVideoTitle As String

    ' .... Class INI
    Public INI As New clsINI


    ' .... Windows offers a way to turn these off, as part of the Structured Exception Handling API.
    ' .... Whilst this API also allows you to intercept any UAE and keep your application running, in this case we don't care
    ' .... about that, since all our code has stopped running and we just want to stop the message showing.
    ' .... You do this by calling the SetErrorMode API call:
    ' .... knowledge base article KBID 309366 > http://support.microsoft.com/default.aspx?scid=kb;en-us;309366

    ' .... Prevent the Crash
    Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

    Private Const SEM_FAILCRITICALERRORS = &H1
    Private Const SEM_NOGPFAULTERRORBOX = &H2
    Private Const SEM_NOOPENFILEERRORBOX = &H8000&

    Private m_bInIDE As Boolean

    ' ... NOW Init the control's XP, Vista and Seven (Vienna)
    Private Type tagInitCommonControlsEx
        lngSize As Long: lngICC As Long
    End Type

    Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
    Private Const ICC_USEREX_CLASSES = &H200

    ' .... Play Sound Resource
    Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
    Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

    Private Const SND_SYNC = &H0
    Private Const SND_ASYNC = &H1
    Private Const SND_NODEFAULT = &H2
    Private Const SND_MEMORY = &H4
    Private Const SND_ALIAS = &H10000
    Private Const SND_FILENAME = &H20000
    Private Const SND_RESOURCE = &H40004
    Private Const SND_ALIAS_ID = &H110000
    Private Const SND_ALIAS_START = 0
    Private Const SND_LOOP = &H8
    Private Const SND_NOSTOP = &H10
    Private Const SND_VALID = &H1F
    Private Const SND_NOWAIT = &H2000
    Private Const SND_VALIDFLAGS = &H17201F
    Private Const SND_RESERVED = &HFF000000
    Private Const SND_TYPE_MASK = &H170007

    Private Const WAVERR_BASE = 32
    Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
    Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
    Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
    Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
    Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)

    Private m_snd() As Byte
    
    ' .... Find and Close the Prev Instance of Application
    Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

    Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Public Const PROCESS_TERMINATE As Long = &H1

    ' .... Constants that are used by the API
    Public Const WM_CLOSE = &H10
    Public Const SYNCHRONIZE = &H100000 ' .... This const is OK?
    Public Const INFINITE = &HFFFFFFFF
    Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

    Public Const GW_HWNDNEXT = 2
    Public mWnd As Long
    
    ' .... Get *flv files or *.* files from Folders and SubFolders
    Public Type FilesCollection
        Count As Long
        Path As New Collection
    End Type

    Public bStop As Boolean
    
    ' .... Browser for Folder's
    Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

    Private Type BROWSEINFO
        hOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
    End Type

    ' .... BrowseInfo Flags
    Public Const BIF_RETURNONLYFSDIRS = &H1
    Public Const BIF_DONTGOBELOWDOMAIN = &H2
    Public Const BIF_STATUSTEXT = &H4
    Public Const BIF_RETURNFSANCESTORS = &H8
    Public Const BIF_EDITBOX = &H10
    Public Const BIF_VALIDATE = &H20
    Public Const BIF_USENEWUI = &H40
    Public Const BIF_NEWDIALOGSTYLE = &H50
    Public Const BIF_BROWSEINCLUDEURLS = &H80
    Public Const BIF_BROWSEFORCOMPUTER = &H1000
    Public Const BIF_BROWSEFORPRINTER = &H2000
    Public Const BIF_BROWSEINCLUDEFILES = &H4000
    Public Const BIF_SHAREABLE = &H8000
Private Sub InitControlsCtx()
 On Local Error GoTo ErrorHandler
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Public Sub Main()
    
    ' .... Init the Controls
    InitControlsCtx
    
    ' .... Show the Form
    Load frmMain: frmMain.Show
End Sub

Public Sub UnloadApp()
    If Not InIDE() Then
        SetErrorMode SEM_NOGPFAULTERRORBOX
    End If
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE()): InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True: IsInIDE = m_bInIDE
End Function

Public Function PlaySoundResource(ByVal SndID As Long) As Long
    Const flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
    On Error GoTo ErrorHandler: DoEvents
    m_snd = LoadResData(SndID, "WAVE")
    PlaySoundResource = PlaySoundData(m_snd(0), 0, flags)
Exit Function
ErrorHandler:
    Err.Clear
End Function

Public Function FileExists(fileName As String) As Boolean
    Dim FSO As New FileSystemObject
    On Local Error GoTo ErrorHandler
    If FSO.FileExists(fileName) Then
        FileExists = True
    Else
        FileExists = False
    End If
    Set FSO = Nothing
Exit Function
ErrorHandler:
    Set FSO = Nothing
        MsgBox "Error: " & Err.Number & vbCr & Err.Description, vbExclamation, App.Title
    Err.Clear
End Function

Public Function FileExtensionFromPath(strPath As String) As String
    On Local Error GoTo ErrorHandler
    FileExtensionFromPath = Right$(strPath, (Len(strPath) - InStrRev(strPath, ".")) + 1)
Exit Function
ErrorHandler:
        FileExtensionFromPath = ""
    Err.Clear
End Function

Public Function ForceClose() As Boolean
    Dim hProcess As Long
    On Local Error GoTo ErrorHandler
    
    If INI.GetKeyValue("PROCESSID", "PID") <> Empty Then _
    PID = INI.GetKeyValue("PROCESSID", "PID") Else PID = 0
    
    If PID = 0 Then
            ForceClose = False
        Exit Function
    Else
        ForceClose = True
    End If
    
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, PID)
    TerminateProcess hProcess, 0&
    
Exit Function
ErrorHandler:
        ForceClose = False
    Err.Clear
End Function

Public Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long: Dim test_pid As Long: Dim test_thread_id As Long
    On Local Error Resume Next
    ' .... Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        ' .... Check if the window isn't a child
        If GetParent(test_hwnd) = 0 Then
            ' .... Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
            If test_pid = target_pid Then
                InstanceToWnd = test_hwnd
                Exit Do
            End If
        End If
        '.... Retrieve the next window
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Public Function WindowToProcessId(ByVal hWnd As Long) As Long
    Dim lpProc As Long: Call GetWindowThreadProcessId(hWnd, lpProc)
    WindowToProcessId = lpProc
End Function

Public Function LongPathName(ByVal fileName As String) As String
    Dim Length As Long, res As String
    On Error Resume Next
    res = String$(MAX_PATH, 0)
    Length = GetLongPathName(fileName, res, Len(res))
    If Length And Err = 0 Then
        LongPathName = Left$(res, Length)
    End If
End Function

Private Sub GetDirs(ByVal sDir As String, cDirAttr As VbFileAttribute, cCol As FilesCollection, Optional sFilter As String = "*.*")
    Dim sStr1 As String
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    sStr1 = Dir$(sDir + sFilter, cDirAttr)
    While sStr1 <> ""
    DoEvents
    If sStr1 <> "." And sStr1 <> ".." Then
        If (GetAttr(sDir & sStr1) And vbDirectory) = vbDirectory Then
            cCol.Path.Add sDir & sStr1
        End If
    End If
        sStr1 = Dir
    Wend
    cCol.Count = cCol.Path.Count
End Sub

Public Sub GetFiles(sDir As String, cFileAttr As VbFileAttribute, cCol As FilesCollection, Optional sFilter As String = "*.flv")
    Dim sStr1 As String
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    sStr1 = Dir$(sDir + sFilter, cFileAttr)
    DoEvents
    If bStop = True Then Exit Sub
    While sStr1 <> ""
        cCol.Path.Add sDir & sStr1
        sStr1 = Dir
    Wend
    cCol.Count = cCol.Path.Count
End Sub

Private Sub GetSubDirs(ByVal sDir As String, cDirAttr As VbFileAttribute, cCol As FilesCollection)
    Dim iTmp As Integer: Dim cCol1 As FilesCollection
    GetDirs sDir, cDirAttr, cCol1
    For iTmp = 1 To cCol1.Count
        cCol.Path.Add cCol1.Path(iTmp)
        GetSubDirs cCol1.Path(iTmp), cDirAttr, cCol
    Next
    cCol.Count = cCol.Path.Count
End Sub

Public Sub GetSubFiles(sDir As String, cDirAttr As VbFileAttribute, cFileAttr As VbFileAttribute, cCol As FilesCollection, Optional sFilter As String = "*.flv")
    Dim iTmp As Integer: Dim sStr1 As String: Dim cCol1 As FilesCollection
    GetSubDirs sDir, cDirAttr, cCol1
    For iTmp = 1 To cCol1.Count
        sStr1 = sStr1 & cCol1.Path(iTmp)
        GetFiles sStr1, cFileAttr, cCol
        sStr1 = Empty
    Next
    GetFiles sDir, cFileAttr, cCol, sFilter
    cCol.Count = cCol.Path.Count
End Sub

Public Function ChkLst(iList As ListBox)
    On Local Error Resume Next
    Dim i As Integer, x As Integer
        For i = 0 To iList.ListCount - 1
            For x = 0 To iList.ListCount - 1
                If (iList.List(i) = iList.List(x)) And x <> i Then
                    iList.RemoveItem (i)
                End If
            Next x
        Next i
End Function

Public Function DialogPrintSetup(hWnd As Long) As Boolean
    Dim x As Long, PD As PRINTDLGSTRUC
    On Local Error GoTo ErrorHandler
        PD.lStructSize = Len(PD): PD.hWnd = hWnd: PD.flags = PD_PRINTSETUP: x = PrintDlg(PD)
    DialogPrintSetup = True
Exit Function
ErrorHandler:
        DialogPrintSetup = False
    Err.Clear
End Function

Public Function GetDefaultPrinter() As String
    Dim x As Long, szTmp As String, dwBuf As Long
    On Error GoTo ErrorHandler
    dwBuf = 1024
    szTmp = Space(dwBuf + 1)
    x = GetProfileString("windows", "device", "", szTmp, dwBuf)
    GetDefaultPrinter = Trim(Left(szTmp, x))
Exit Function
ErrorHandler:
        GetDefaultPrinter = "Error# " & Err.Number
    Err.Clear
End Function

Public Function ResetDefaultPrinter(szBuf As String) As Boolean
    Dim x As Long
    On Local Error GoTo ErrorHandler
    x = WriteProfileString("windows", "device", szBuf)
    x = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
    ResetDefaultPrinter = True
Exit Function
ErrorHandler:
        ResetDefaultPrinter = False
    Err.Clear
End Function

Public Function SetDefaultPrinter(objPrn As Printer) As Boolean
    Dim x As Long, szTmp As String
    On Local Error GoTo ErrorHandler
    szTmp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    x = WriteProfileString("windows", "device", szTmp)
    x = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
    SetDefaultPrinter = True
Exit Function
ErrorHandler:
        SetDefaultPrinter = False
    Err.Clear
End Function

Public Function DialogConnectToPrinter() As Boolean
    On Local Error GoTo ErrorHandler
        Shell "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus
    DialogConnectToPrinter = True
Exit Function
ErrorHandler:
        DialogConnectToPrinter = False
    Err.Clear
End Function

Public Function ChangeTextLong(Text As String, HowLong As Double) As String
    Dim s() As String, i As Integer, newS() As String, j As Integer, u As Integer, TextToCheck As String
    Dim n As Integer, Y As Double, sp As Double
    s = Split(Text, vbCrLf)
    ReDim newS(0)
    Printer.ScaleMode = 6
    For i = 0 To UBound(s)
        If Printer.TextWidth(s(i)) > HowLong Then
             TextToCheck = s(i)
startCheck:
u = UBound(newS) + 1
             j = Len(TextToCheck)
             Do
                 j = j - 1
                 ReDim Preserve newS(u)
                 newS(u) = Mid(TextToCheck, 1, j)
                 If Printer.TextWidth(newS(u)) <= HowLong Then Exit Do
             Loop
             If Printer.TextWidth(Mid(TextToCheck, j + 1, Len(TextToCheck) - j)) > HowLong Then
                    TextToCheck = Mid(TextToCheck, j + 1, Len(TextToCheck) - j)
                 GoTo startCheck
             Else
                 u = UBound(newS) + 1
                    ReDim Preserve newS(u)
                 newS(u) = Mid(TextToCheck, j + 1, Len(TextToCheck) - j + 1)
             End If
        Else
            u = UBound(newS) + 1
                    ReDim Preserve newS(u)
            newS(u) = s(i)
        End If
    Next i
    For i = 1 To UBound(newS)
        If i = 1 Then
            ChangeTextLong = newS(i)
        Else
            ChangeTextLong = ChangeTextLong & vbCrLf & newS(i)
        End If
    Next i
    
    Y = 10: sp = 4 ' .... Spessore delle righe
    s = Split(ChangeTextLong(Text, 74), vbCrLf)
    For n = 0 To UBound(s)
        ' .... Stampo le righe
        Printer.CurrentX = 10
        Printer.CurrentY = Y + sp / 2 - Printer.TextHeight(s(n)) / 2
        Printer.Print s(n)
        Y = Y + sp
    Next n
    
End Function

Public Function FolderBrowse(hWnd As Long, szDialogTitle As String) As String
    Dim x As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
    BI.hOwner = hWnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If x Then
        wPos = InStr(szPath, Chr(0))
        FolderBrowse = Left$(szPath, wPos - 1)
    Else
        FolderBrowse = ""
    End If
Exit Function
ErrorHandler:
    Err.Clear
End Function
