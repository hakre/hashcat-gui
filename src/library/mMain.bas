Attribute VB_Name = "mMain"

'
' API
'

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
    
Private Declare Function GetModuleHandle Lib "kernel32" _
    Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
    
Private Declare Function GetCurrentProcess Lib "kernel32" _
    () As Long

Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long



Public HCGUI_Inifile As String
Public HCGUI_Mainform As fMain
Public HCGUI_BinFile As String
Public HCGUI_BinOs As eAcBinOs

Public Enum eAcBinOs
    Windows = 0
    Wine = 1
End Enum

Public Enum eAcTriggerEvents
    ChangeEvent = 1
    ClickEvent = 2
End Enum


Public Function Cnv_Str2Dec(ByVal str As String) As Variant
Dim i As Long, l As Long, tstr As String, c As Long

    str = Trim(str)
    
    ' filter string to decimals (this is not round for a good reason)
    l = Len(str)
    If l > 0 Then
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If (c = 48 And Len(tstr) > 0) Or (c > 48 And c < 58) Then
                tstr = tstr + Chr(c)
            End If
        Next i
        str = tstr
    End If

    If str = "" Then
        Cnv_Str2Dec = ""
    Else
        If Len(str) > 29 Then
            str = "79228162514264337593543950335"
        End If
        If str > "79228162514264337593543950335" Then
            str = "79228162514264337593543950335"
        End If
        Cnv_Str2Dec = CDec(str)
    End If

End Function
'select hashcat exe by a dialog
'returns an empty string if cancel / failed
Public Function HCGUI_bin_askfor(Optional ByVal hWnd As Long = 0, Optional BinOs As eAcBinOs = -1, Optional sInitFile As String = "") As String
Dim cc As cCommonDialog
Dim sFile As String
Dim sFilter As String
Dim sFilterWin As String
Dim sFilterLinux As String
Dim sInitDir As String
Dim cFile As New cFileinfo

    ' Filter for older, old and 32/64 bit hashcat cli binaries
    sFilterWin = "hashcat.exe;hashcat-cli.exe;hashcat-cli32.exe;hashcat-cli64.exe"
    sFilterLinux = Replace(sFilterWin, ".exe", ".bin")

    ' WINE compat layer
    If BinOs = -1 Then BinOs = HCGUI_BinOs
    If BinOs = Windows Then
        sFilter = "hashcat Windows Executeable|" & sFilterWin & "|hashcat Linux Executeable|" & sFilterLinux
    Else
        sFilter = "hashcat Linux Executeable|" & sFilterLinux & "|hashcat Windows Executeable|" & sFilterWin
    End If
    
    ' initial directory
    cFile.Path = sInitFile
    sInitDir = cFile.ExistingDir
    
    ' display dialog
    Set cc = New cCommonDialog
    
    If cc.VBGetOpenFileName(sFile, , , , , , _
        sFilter & "|Executeables|*.bin;*.exe|All Files (*.*)|*.*", , sInitDir, "Select hashcat Executeable", "exe", hWnd, OFN_HideReadOnly) Then
        HCGUI_bin_askfor = sFile
    End If
    
End Function
'
' default name of the binary
'
' returns an existing file or empty string in case there is no way to determine
' the default hashcat binary
'
Public Function HCGUI_bin_default() As String
Dim sDefBasename As String
Dim oFi As New cFileinfo
Dim sDefault As String
Dim oFileDefault As New cFileinfo


    'default binary basename
    sDefBasename = "hashcat-cli32.exe"
    If HCGUI_is64bit() Then
        sDefBasename = "hashcat-cli64.exe"
    End If
    
    oFi.Path = HCGUI_directory(0)
    
    oFileDefault.Path = oFi.Dirname & sDefBasename
    
    If oFileDefault.Exists Then
        HCGUI_bin_default = oFileDefault.Path
    Else
        HCGUI_bin_default = ""
    End If

End Function
'
'
' guess the name of the binary
'
' returns the existing file or an empty string in case there is no way to determine
' a working version
'
Public Function HCGUI_bin_guess(Optional bInteractive As Boolean = False, Optional hWnd As Long = 0, Optional BinOs As eAcBinOs = -1) As String
Dim oFi As New cFileinfo
Dim sFile As String
Dim sDefBasename As String
Dim sDefault As String

    'first check the ini
    sFile = INIRead("default", "hashcat", HCGUI_Inifile)
    
    'default binary basename
    sDefBasename = "hashcat-cli32.exe"
    If Environ("PROCESSOR_ARCHITECTURE") = "AMD64" Then
        sDefBasename = "hashcat-cli64.exe"
    End If
    
    If BinOs = -1 Then BinOs = HCGUI_BinOs
    If BinOs = Wine Then
        ' FIXME it's just the same, fix this (2010-12-04)
        sDefBasename = sDefBasename
    End If
            
    'create the default to be able to compare
    oFi.Path = HCGUI_directory(0)
    sDefault = oFi.Dirname & sDefBasename
    If Not oFi.Exists Then
        sDefault = ""
    Else
        If BinOs = Wine Then
            'pass basename only on wine sothat exec works hopefully properly
            sDefault = sDefBasename
        End If
    End If

    'if there is no file, ask for it if allowed to be interactive
    If sFile = "" Then
        If sDefault = "" And bInteractive Then
            sDefault = HCGUI_bin_askfor(hWnd)
        End If
        sFile = sDefault
    End If
    
    HCGUI_bin_guess = sFile
End Function
'check if a name of an hashcat binary is valid
Public Function HCGUI_bin_isvalid(Name As String) As Boolean

    HCGUI_bin_isvalid = True
    
End Function


'
' get version string of hashcat binary
'
' will return "Error" in case of an error
'
Public Function HCGUI_binver(Optional Path As String = "")
Dim sOut As String

    If Path = "" Then
        Path = HCGUI_bin().Path
    End If
    
    sOut = GetCommandOutput(Path & " --version")
    
    If sOut = "" Or Len(sOut) > 8 Then
        sOut = "Error"
    Else
        sOut = Replace(sOut, vbCrLf, "")
        sOut = Replace(sOut, vbCr, "")
    End If
    
    
        
    HCGUI_binver = sOut

End Function


' guess 64 bit state
Public Function HCGUI_is64bit()
Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    HCGUI_is64bit = bolFunc

End Function

' get the right-left corner of a control (as pointapi in pixels) on screen
Public Function POS_control_bottomright(aControl As Control) As POINTAPI
Dim p1 As POINTAPI
Dim r As Long
Dim Scalemode As Long
Dim aForm As Form

    Set aForm = aControl.Parent

    Scalemode = aForm.Scalemode
    aForm.Scalemode = vbPixels
    With aControl
        p1.x = .Left + .Width
        p1.y = .Top + .Height
    End With
    r = ClientToScreen(aForm.hWnd, p1)
    aForm.Scalemode = Scalemode
        
POS_control_bottomright = p1
End Function

' get the top-left corner of a control (as pointapi in pixels) on screen
Public Function POS_control_topleft(aControl As Control) As POINTAPI
Dim p1 As POINTAPI
Dim r As Long
Dim Scalemode As Long
Dim aForm As Form

    Set aForm = aControl.Parent
        
    Scalemode = aForm.Scalemode
    aForm.Scalemode = vbPixels
    With aControl
        p1.x = .Left
        p1.y = .Top
    End With
    r = ClientToScreen(aForm.hWnd, p1)
    aForm.Scalemode = Scalemode
        
POS_control_topleft = p1
End Function
'get a directory (useable to attach a file to -> \ at the end is there)
'
' vIndex:
'    0 - applications working directory
'    1 - applications data directory (ini)
'    100 - data dir (dicts, rules)
'    101 - data:dictionary dir
'    102 - data:rules dir
Public Function HCGUI_directory(Optional vIndex As Variant = 0) As String
Dim sPath As String
Dim oFi As cFileinfo

    Select Case vIndex
        Case 0: 'application directory
            Set oFi = New cFileinfo
            oFi.Path = App.Path
            HCGUI_directory = oFi.Dirname
            
        Case 1: 'appdata dir
            HCGUI_directory = HCGUI_directory(1026) & "hashcat\"
            
        Case 100: 'data dir (user)
            HCGUI_directory = HCGUI_directory(1005) & "hashcat\"
            
        Case 101: 'data:dictionaries dir
            HCGUI_directory = HCGUI_directory(100) + "dicts\"
            
        Case 102: 'data:rules dir
            HCGUI_directory = HCGUI_directory(100) + "rules\"
            
        Case 103: 'data:salts dir
            HCGUI_directory = HCGUI_directory(100) + "salts\"
            
        Case 1005, 1026: 'shell: 1005 = my documents, 1026 = appdata
            Set oFi = New cFileinfo
            oFi.Path = WINAPI_GetSpecialFolderLocation(vIndex - 1000)
            HCGUI_directory = oFi.Dirname
            
        Case Else:
            Stop
            
    End Select
    
End Function


Public Function HCGUI_Plains_Expand(oPlains As cPlains) As cPlains
Dim oPlain As cPlainfile
Dim newPlains As New cPlains
Dim newPlain As cPlainfile
Dim oFile As cFileinfo
    
    Set colNew = New Collection
        
    For Each oPlain In oPlains
        If oPlain.Checked Then
            'we are only interested in checked plains
            
            If Not oPlain.Fileinfo.Exists Then
                'skip non existant files
                ' -- nothing to do --
            ElseIf oPlain.Fileinfo.isDir Then
                'directories need to be extended first
                For Each oFile In oPlain.Fileinfo.Files
                    If oFile.Exists And oFile.isFile Then
                        Set newPlain = New cPlainfile
                        newPlain.FileName = oFile.Path
                        newPlain.Checked = True
                        Call newPlains.Add(newPlain)
                    End If
                Next
            Else
                'just files
                Set newPlain = New cPlainfile
                newPlain.FileName = oPlain.FileName
                newPlain.Checked = True
                Call newPlains.Add(newPlain)
            End If
        End If
    Next

'Dim sortPlain As cPlainfile

Set HCGUI_Plains_Expand = newPlains
End Function
Public Sub HCGUI_recent_from_ini(oRecent As cRecent, sFile As String, Optional sPrefix As String = "", Optional sSection As String = "recent")
Dim iCount As Long
Dim i As Long

    oRecent.Clear
    
    iCount = INIReadInt(sSection, sPrefix & "count", sFile)
    
    If iCount > 0 Then
        For i = iCount To 1 Step -1
            oRecent.Touch INIRead(sSection, sPrefix & "item" + CStr(i), sFile)
        Next i
    End If

End Sub

Public Function HCGUI_recent_to_ini(oRecent As cRecent, sFile As String, Optional sPrefix As String = "", Optional sSection As String = "recent")
Dim i As Long

    Call INIWrite(oRecent.Count, sSection, sPrefix & "count", sFile)
    
    If oRecent.Count > 0 Then
        For i = 1 To oRecent.Count
            Call INIWrite(oRecent.Item(i).Path, sSection, sPrefix & "item" & CStr(i), sFile)
        Next
    End If

End Function

Public Function MinMax(ByVal Value As Long, Min As Long, Max As Long) As Long

    If Value < Min Then Value = Min
    If Value > Max Then Value = Max

MinMax = Value
End Function

Public Function newLvPlains(list As ListView, Helper As cLvHelper, Form As fMain) As cLvPlains
Dim oLvPlains As cLvPlains

    Set oLvPlains = New cLvPlains
    Call oLvPlains.Init(list, Helper, Form)
    
Set newLvPlains = oLvPlains
End Function

Public Function newTwHelper(cMainWin As fMain, cToolWin As Form) As cTwHelper
Dim oNew As New cTwHelper

    Call oNew.Init(cMainWin, cToolWin)
    
Set newTwHelper = oNew
End Function


'of a window is (at least one pixel) out of the screen, it
'will be moved inside it.
'
'returns: number of movements made (max 2), 0 if no move has occured
Public Function POS_window_forceonscreen(aForm As Form) As Long
Dim iMoved As Long
            
    With aForm
        'X: min X
        If .Left < 0 Then
            .Left = 0
            iMoved = iMoved + 1
        End If
        'X: max X
        If .Left + .Width > Screen.Width Then
            .Left = Screen.Width - .Width
            iMoved = iMoved + 1
        End If
        'Y: min Y
        If .Top < 0 Then
            .Top = 0
            iMoved = iMoved + 1
        End If
        'Y: max Y
        If .Top + .Height > Screen.Height Then
            .Top = Screen.Height - .Height
            iMoved = iMoved + 1
        End If
    End With
    
POS_window_forceonscreen = iMoved
End Function

Public Sub Sleep(sleepTime As Single)
Dim t As Single
t = Timer
Do
    DoEvents
    
    If t > Timer Then Exit Do 'preserve midnight probs
    
Loop Until (t + sleepTime >= Timer)


End Sub

Public Function file_get_contents(sFile As String) As String
Dim iFp As Long
Dim sContents As String


    iFp = FreeFile
    
    Open sFile For Input As #iFp
    'While Not EOF(iFp)
    'Wend
        'exit while
    sContents = Input(LOF(iFp), iFp)
    
    Close iFp
    Close iFp
    Close iFp
    Close iFp
    

file_get_contents = sContents
End Function



'
Public Sub Main()
Dim oFi As New cFileinfo
Dim cCmd As New cCommandline
Dim sTest As String

    'validate directories
    oFi.Path = HCGUI_directory(1)
    If Not oFi.Exists Then
        MsgBox "hashcat appdata directory does not exists: " & vbCrLf & vbCrLf & oFi.Path & vbCrLf & vbCrLf & "Ensure that his directory exists prior to start hashcat-gui.", vbCritical, "Directory Validation"
        Exit Sub
    End If
       

    'init inifilename
    HCGUI_Inifile = HCGUI_directory(1) & "hashcat-gui.ini"
    
    'init os
    HCGUI_BinOs = MinMax(INIReadInt("default", "wineenabled", HCGUI_Inifile), 0, 1)
        
    'init binfile
    HCGUI_BinFile = "*"
    Call HCGUI_bin(False) 'read from ini or create by default
    
    
    Set HCGUI_Mainform = New fMain
    
    If cCmd.existsPassedFile Then
        Call HCGUI_Mainform.FileOpen(cCmd.passedFile)
    End If
    
    HCGUI_Mainform.Show

End Sub
' set the textual represenation of a shortcut to a certain
' menu entry.
Public Sub Menu_SetShortcut(oMenu As Menu, sText As String)
Dim sCaption As String
Dim sShort As String
Dim p As Long

    sCaption = oMenu.Caption
    
    'some safe things
    If sCaption = "-" Then
        MsgBox "Internal Error: Menu [" & oMenu.Name & "] not applicable for shortcut [" & sText & "]"
    ElseIf InStr(sCaption, vbTab) <> 0 Then
        p = InStr(sCaption, vbTab)
        sShort = Mid(sCaption, p + 1)
        sCaption = Left(sCaption, p - 1)
        If sText = "" Then
            Stop
        Else
            Stop
        End If
    Else
        sCaption = sCaption + vbTab + sText
    End If
    
    oMenu.Caption = sCaption
        
End Sub
'
' oledd_is_files
'
' @since  0.1
' @return boolena true if data contains files, false if not
'
Public Function oledd_is_files(Data As Object)
    Dim retval As Boolean
    
    retval = False

    If Data.GetFormat(vbCFFiles) Then
        If Data.Files.Count > 0 Then
            retval = True
        End If
    End If
    
    oledd_is_files = retval
    
End Function
'
' make a parameter safe to be passed as one
'
Public Function paramsafe(ByVal Param As String, BinOs As eAcBinOs) As String
        
    'convert path seperators on wine
    If BinOs = Wine Then
        Param = Replace(Param, "\", "/")
    End If
        
    'handle spaces (both os'es)
    If InStr(Param, " ") Then
        Param = """" & Param & """"
    End If

    paramsafe = Param
End Function

'
' textbox_select_all
'
' select all text of a textbox or similar control
'
' @since  0.1
' @return void
'
Public Sub textbox_select_all(Control As Object)
    
    Control.SelStart = 0
    Control.SelLength = Len(Control.Text)

End Sub


Public Function HCGUI_bin(Optional bInteractive As Boolean = False, Optional hWnd As Long = 0) As cFileinfo
Dim oFi As New cFileinfo
Dim sFile As String

    ' bin is saved in the ini
    If HCGUI_BinFile = "*" Then
       HCGUI_BinFile = INIRead("default", "hashcat", HCGUI_Inifile)
       If HCGUI_BinFile = "" Then
            HCGUI_BinFile = HCGUI_bin_guess(bInteractive, hWnd)
       End If
    End If
    
    oFi.Path = HCGUI_BinFile
    
    Set HCGUI_bin = oFi
End Function


Public Function HCGUI_job_from_ini(sFile As String, Optional sSection As String = "session") As cJob
Dim oJob As New cJob
Dim iNum As Long
Dim i As Long

    'attack
    oJob.RecoveryMode = INIReadInt(sSection, "attackmode", sFile)

    'hash
    oJob.HashFile = INIRead(sSection, "hash", sFile)
    oJob.HashMode = INIReadInt(sSection, "hashmode", sFile)
    oJob.hashSeperator = INIRead(sSection, "hashseperator", sFile)
    oJob.hashRemove = CBool(INIReadInt(sSection, "hashremove", sFile))
    
    ' bruteforce options
    oJob.BruteChars = INIRead(sSection, "brutechars", sFile)
    oJob.BruteLen.Value = INIRead(sSection, "brutelen", sFile)
    
    'limit
    oJob.Limit = Cnv_Str2Dec(INIRead(sSection, "limit", sFile))
    
    'outfile
    oJob.OutFile.External = INIRead(sSection, "out", sFile)
    
    'plains
    iNum = INIReadInt(sSection, "plainnum", sFile)
    If iNum > 0 Then
        For i = 1 To iNum
            Call oJob.Plains.Import(INIRead(sSection, "plain" & CStr(i), sFile))
        Next i
    End If
    
    'rules
    oJob.RuleCount = INIRead(sSection, "rulecount", sFile)
    oJob.RuleFile = INIRead(sSection, "rulefile", sFile)
    oJob.RuleMode = INIReadInt(sSection, "rulemode", sFile)
    
    'saltfile
    oJob.SaltFile.External = INIRead(sSection, "saltfile", sFile)
    
    'segment
    oJob.Segment = INIReadInt(sSection, "segment", sFile)
    
    'skip
    oJob.Skip = Cnv_Str2Dec(INIRead(sSection, "skip", sFile))
    
    'threads
    iNum = INIReadInt(sSection, "threads", sFile)
    If iNum < 1 Then
        iNum = Val(Environ("NUMBER_OF_PROCESSORS"))
    End If
    oJob.Threads = iNum
    
    'toggle
    oJob.ToggleLen.Value = INIRead(sSection, "togglelen", sFile)
    
    Set HCGUI_job_from_ini = oJob
    
End Function

Public Function HCGUI_job_to_ini(oJob As cJob, sFile As String, Optional sSection As String = "session") As Long
Dim r As Long
Dim i As Long
Dim oPlain As cPlainfile

    ' attackmode
    r = INIWrite(oJob.RecoveryMode, sSection, "attackmode", sFile)
    
    ' hash
    r = INIWrite(oJob.HashFile, sSection, "hash", sFile)
    r = INIWrite(oJob.HashMode, sSection, "hashmode", sFile)
    r = INIWrite(oJob.hashSeperator, sSection, "hashseperator", sFile)
    r = INIWrite(CInt(oJob.hashRemove), sSection, "hashremove", sFile)
    
    
    ' bruteforce options
    r = INIWrite(oJob.BruteChars, sSection, "brutechars", sFile)
    r = INIWrite(oJob.BruteLen.Value, sSection, "brutelen", sFile)
    
    'limit
    r = INIWrite(oJob.Limit, sSection, "limit", sFile)
    
    'outfile
    r = INIWrite(oJob.OutFile.External, sSection, "out", sFile)
    
    'plains
    i = 0
    For Each oPlain In oJob.Plains
        i = i + 1
        r = INIWrite(oPlain.External, sSection, "plain" & CStr(i), sFile)
    Next
    r = INIWrite(oJob.Plains.Count, sSection, "plainnum", sFile)
    
    'rules
    r = INIWrite(oJob.RuleCount, sSection, "rulecount", sFile)
    r = INIWrite(oJob.RuleFile, sSection, "rulefile", sFile)
    r = INIWrite(oJob.RuleMode, sSection, "rulemode", sFile)
    
    'saltfile
    r = INIWrite(oJob.SaltFile.External, sSection, "saltfile", sFile)
    
    'segment
    r = INIWrite(oJob.Segment, sSection, "segment", sFile)
    
    'skip
    r = INIWrite(oJob.Skip, sSection, "skip", sFile)
    
    'threads
    r = INIWrite(oJob.Threads, sSection, "threads", sFile)
    
    'toggle
    r = INIWrite(oJob.ToggleLen.Value, sSection, "togglelen", sFile)
    
    
End Function

Public Function HCGUI_Form_IsLoaded(ConcreteForm As Form) As Boolean
Dim iterateForm As Object

    If Not ConcreteForm Is Nothing Then
        For Each iterateForm In Forms
            If iterateForm Is ConcreteForm Then
                HCGUI_Form_IsLoaded = True
                Exit Function
            End If
        Next
    End If

End Function
