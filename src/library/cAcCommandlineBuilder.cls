VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAcCommandlineBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Commandline Builder
' needs a fMain-form to gain additional values from
'

Private m_Display As Long
Private m_Binary As String
Private m_Prefix As String
Private m_Suffix As String
Private m_hasPrefix  As Boolean
Private m_hasFullpath As Boolean
Private m_hasFullpathParams As Boolean
Private m_form As fMain
Private m_set_EscapeBinaryWin As Boolean   'FIXME not used
Private m_set_EscapeBinaryLinux As Boolean 'FIXME not used
Private m_set_SpaceInParamEncloseParam As Boolean
Private m_set_SpaceInParamEscapeSpace As Boolean
Private m_set_WinToLinuxSlashes As Boolean
Private m_set_ParamEscapeWin As Boolean

' NumMin(sParams, m_form.skipText(0).Text, "-s", 0)
Private Function DecMin(sParams As String, ByVal sData As String, sFlag As String, iDefault As Long, Optional iMin As Long = 0)
Dim dData As Variant

    dData = Cnv_Str2Dec(sData)
        
    If TypeName(dData) = "Decimal" Then
        If dData < iMin Then dData = iMin
        If dData > CDec("18446744073709551615") Then dData = CDec("18446744073709551615")
    Else
        dData = CDec(0)
    End If
    
    
    If dData >= iMin And dData <> iDefault Then
        sParams = sParams & " " & sFlag & " " & CStr(dData)
    End If
    
DecMin = sParams
End Function

'
Private Function NumMin(sParams As String, ByVal sData As String, sFlag As String, iDefault As Long, Optional iMin As Long = 0)
Dim iValue As Long
Dim dValue As Double
Dim vValue As Variant
Dim iLimitPlaces As Long

    sData = Trim(sData)

    iLimitPlaces = 15
    If Len(sData) > iLimitPlaces Then
        sData = Left(sData, iLimitPlaces)
    End If

    
    dValue = Val(sData)
    vValue = Round(dValue, 0)
    'Debug.Print "Value (" & sData & ") Rounded is: " & vValue
            
    If vValue >= iMin And vValue <> iDefault Then
        sParams = sParams & " " & sFlag & " " & CStr(vValue)
    End If
        
NumMin = sParams
End Function

Public Property Get Binary() As String
    Binary = m_Binary
End Property
Public Property Let Binary(Path As String)
    m_Binary = Path
End Property
' returns empty commandline if form is not attached
Public Property Get Commandline()
If m_form Is Nothing Then
    Exit Property
End If

Dim sLine As String
Dim oFi As New cFileinfo

    'debug
    If False Then
        sLine = "[" & CStr(Me.Display) & "] "
    End If
    
    'prefix
    If m_hasPrefix Then
        sLine = sLine & m_Prefix
    End If
    
    'binary
    If m_hasFullpath Then
        sLine = sLine & EscapeParam(m_Binary)
    Else
        oFi.Path = m_Binary
        sLine = sLine & EscapeParam(oFi.Basename)
    End If
    
    'parameters
    sLine = sLine & " " & Me.Parameters
    
    'suffix
    sLine = sLine & m_Suffix
    
Commandline = Me.Prefix & sLine
End Property
Public Property Get Display() As Long
    Display = m_Display
End Property
Public Property Let Display(Mode As Long)
    Mode = MinMax(Mode, 0, 2)
    m_Display = Mode
    
    m_hasFullpath = False
    m_hasFullpathParams = False
    
    m_set_SpaceInParamEncloseParam = False
    m_set_SpaceInParamEscapeSpace = False
    m_set_WinToLinuxSlashes = False
    m_set_ParamEscapeWin = False
    m_set_EscapeBinaryWin = False
    m_set_EscapeBinaryLinux = False
    
    Select Case Mode
        Case 0: 'win
            m_set_EscapeBinaryWin = True
            m_set_SpaceInParamEncloseParam = True
            m_set_ParamEscapeWin = True
        
        Case 1: 'linux
            m_set_SpaceInParamEscapeSpace = True
            m_set_WinToLinuxSlashes = True
            m_set_EscapeBinaryLinux = True
        
        Case 2: 'wine
            m_hasFullpath = True
            m_hasFullpathParams = True
            
            m_set_EscapeBinaryWin = True
            m_set_SpaceInParamEncloseParam = True
            m_set_ParamEscapeWin = True
            
    End Select

End Property
Public Property Get Form() As fMain
    Set Form = m_form
End Property
Public Property Set Form(oForm As fMain)
    Set m_form = oForm
    Me.Display = oForm.commandlineDisplay
    Me.Binary = oForm.Binary
End Property
Public Property Get hasFullpath() As Boolean
    hasFullpath = m_hasFullpath
End Property
Public Property Let hasFullpath(Value As Boolean)
    m_hasFullpath = Value
End Property
Public Property Get hasFullpathParams() As Boolean
    hasFullpath = m_hasFullpathParams
End Property
Public Property Let hasFullpathParams(Value As Boolean)
    m_hasFullpathParams = Value
End Property
'
' Internet Resources related to the topic:
'
' - http://ss64.com/nt/syntax-esc.html
' - http://ss64.com/bash/syntax-quoting.html
'
Public Function EscapeParam(ByVal ParamRaw As String)
Dim bSpaceInParam As Boolean
Dim Param As String

    'copy raw param so that it can stay untouched
    Param = ParamRaw

    'is there a space in the parameter?
    If InStr(ParamRaw, " ") Then
        bSpaceInParam = True
    End If
    
    'windows to linux slashes converstion
    If m_set_WinToLinuxSlashes Then
         Param = Replace(Param, "\", "/")
    End If
    
    'deal with windows shell characters in the param
    If m_set_ParamEscapeWin Then
        Param = zEscapeWinChars(Param)
    End If
    
    'deal with spaces in the param
    If bSpaceInParam Then
        'escape those
        If m_set_SpaceInParamEscapeSpace Then
            Param = Replace(Param, " ", "\ ")
        End If
        'enclose whole param
        If m_set_SpaceInParamEncloseParam Then
            Param = """" & Param & """"
        End If
    End If
    
EscapeParam = Param
End Function
Private Function HandleRelativity(Path As String) As String

    HandleRelativity = Path
    If m_hasFullpathParams Then Exit Function

Dim oFile As New cFileinfo
    
    oFile.Path = Path
    HandleRelativity = oFile.RelPath(Me.Binary)

End Function

Public Property Get Parameters() As String
Dim sParams As String
Dim iRecoverymode As Long
Dim sValue As String
Dim iValue As String
Dim sSeperator As String
    
    'attack mode -a
    iRecoverymode = m_form.recoveryModeCombo.ListIndex
    sParams = sParams & "-a " & CStr(iRecoverymode)
    
    
    Select Case iRecoverymode
        Case 0, 1, 2: 'rules / hybrid or combination
        
            Select Case m_form.RuleMode
                Case 1: 'ruleset file -r / --rules-file
                    sValue = m_form.ruleFileText
                    If Len(sValue) Then
                        sParams = sParams & " -r " & EscapeParam(HandleRelativity(sValue))
                    End If
                    
                Case 2: 'ruleset generator -g / --generate-fules
                    sParams = sParams & " -g " & m_form.RuleGenerate
            End Select
            
            If iRecoverymode = 2 Then 'toggle attack
                sParams = NumMin(sParams, m_form.toggleLenText(0).Text, "--toggle-min", 1, 1)
                sParams = NumMin(sParams, m_form.toggleLenText(1).Text, "--toggle-max", 16, 1)
            End If
            
        Case 3: 'bruteforce
        
            sParams = sParams & " --bf-cs-buf " & EscapeParam(m_form.bruteCharsText)
            sParams = NumMin(sParams, m_form.bruteLenText(0).Text, "--bf-pw-min", 1, 1)
            sParams = NumMin(sParams, m_form.bruteLenText(1).Text, "--bf-pw-max", 8, 1)
            
        Case 4: 'permutation (re-use of bruteLenText-boxes)
            sParams = NumMin(sParams, m_form.bruteLenText(0).Text, "--perm-min", 1, 1)
            sParams = NumMin(sParams, m_form.bruteLenText(1).Text, "--perm-max", 8, 1)
            
    End Select
    
    'hash mode -m
    sParams = sParams & " -m " & m_form.HashMode
    
    'saltfile (if applicable)
    If m_form.HashModeHasSalt Then
        If m_form.saltFileCheck.Value Then
            sValue = m_form.saltFileText
            sParams = sParams & " -e " & EscapeParam(HandleRelativity(sValue))
        End If
    End If
    
    'output file -o
    If m_form.outFileCheck.Value Then
        sValue = m_form.outFileText
        sParams = sParams & " -o " & EscapeParam(HandleRelativity(sValue))
    End If
    
    'output format --output-format
    sParams = NumMin(sParams, m_form.outFormatText.ListIndex, "--output-format", 0)
    
    'resources
    sParams = NumMin(sParams, m_form.threadsText.Text, "-n", 8)
    
    'segment size
    sParams = NumMin(sParams, m_form.segmentText.Text, "-c", 32)
    
    'skip
    sParams = DecMin(sParams, m_form.skipText(0).Text, "-s", 0)

    'limit
    sParams = DecMin(sParams, m_form.skipText(1).Text, "-l", 0)

    'hashseperator
    sSeperator = m_form.hashSeperator.Text
    If sSeperator <> ":" And Len(sSeperator) Then
        sParams = sParams & " -p " & EscapeParam(sSeperator)
    End If
    
    'remove found hashes --remove
    If m_form.hashRemove.Value Then
        sParams = sParams & " --remove"
    End If

    'hashes
    sValue = m_form.hashFileText.Text
    If Len(sValue) Then
        sParams = sParams & " " & EscapeParam(HandleRelativity(sValue))
    End If

    'plains
    If Len(Me.plainsList) Then
        sParams = sParams & " " & Me.plainsList
    End If

    Parameters = sParams

End Property

Property Get plainsList() As String
Dim cFile As New cFileinfo
Dim sKey As String
Dim oLi As ListItem
Dim sList As String
Dim sEntry As String
Dim iMouseOld As Long


    iMouseOld = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    'we need a form to gather all files
    If m_form Is Nothing Then Exit Property
    
    For Each oLi In m_form.plainsList.ListItems
        If oLi.Checked Then
            cFile.Path = oLi.ToolTipText
            If cFile.Exists Then
                sEntry = EscapeParam(HandleRelativity(cFile.Path))
                If Len(sEntry) Then
                    sList = sList & IIf(Len(sList), " ", "") & sEntry
                End If
            End If
        End If
    Next
    
    plainsList = sList
    
    Screen.MousePointer = iMouseOld
    
End Property

Public Property Get Prefix() As String
    Prefix = m_Prefix
End Property
Public Property Let Prefix(Prefix As String)
    m_Prefix = Prefix
End Property
Public Property Get hasPrefix() As Boolean
    hasPrefix = m_hasPrefix
End Property
Public Property Let hasPrefix(Value As Boolean)
    m_hasPrefix = Value
End Property
Public Property Let Suffix(sSuffix As String)
    m_Suffix = sSuffix
End Property

Public Property Get Suffix() As String
    Suffix = m_Suffix
End Property

Private Function zEscape(ByVal sString As String, ByVal sEscapeChar As String, sToEscape As String) As String
Dim i As Long, m As Long
Dim sChar As String
Dim sCmd As String


    m = Len(sToEscape)
    If m < 1 Then
        Exit Function
    End If
    
    sCmd = sString
    
    For i = 1 To m
        sChar = Mid(sToEscape, i, 1)
        sCmd = Replace(sCmd, sChar, sEscapeChar & sChar)
    Next i

zEscape = sCmd
End Function

' the following chars \ & | > < ^ must be escaped with ^
Private Function zEscapeWinChars(ByVal sParam As String) As String
Dim aChars() As String
Dim sChar As Variant
Dim c As Long

    'command characters
    aChars = Split("^ & | > <", " ")
           
    For Each sChar In aChars
        sParam = Replace(sParam, sChar, "^" & sChar)
    Next
    
    'double quote(s) - postfix param with a double quote if an
    'odd number of quotes is in there.
    c = COL_stringCount(sParam, """")
    If c Mod 2 = 1 Then
        sParam = sParam & """"
    End If

zEscapeWinChars = sParam
End Function
