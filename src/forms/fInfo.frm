VERSION 5.00
Begin VB.Form fInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About hashcat"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox helpText 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   5
      Text            =   "fInfo.frx":000C
      Top             =   1320
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Frame closeFrame 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   4695
      Begin VB.CommandButton closeCommand 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   375
         Left            =   3300
         TabIndex        =   4
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame infoImageFrame 
      Height          =   1095
      Left            =   -180
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      Begin VB.Label infoCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Bitstream Vera Sans"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -1500
         TabIndex        =   6
         Top             =   420
         Width           =   5910
      End
      Begin VB.Image infoImage 
         Height          =   1395
         Left            =   0
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Image wineImage 
      Height          =   1500
      Left            =   3540
      Picture         =   "fInfo.frx":0017
      ToolTipText     =   "This software has been designed with WINE in mind."
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   1260
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      Height          =   3135
      Left            =   60
      TabIndex        =   1
      Top             =   1860
      Width           =   3315
   End
End
Attribute VB_Name = "fInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' API
'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26

Private Const STYLE_NORMAL = 11

'
' http://groups.google.com/group/microsoft.public.vb.general.discussion/browse_thread/thread/ea0ba7fccd937652
'

Private Const IDC_HAND = 32649&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long


'
' VB
'

Private m_Parent As fMain

Private Function Hashcat_Version() As String
Dim sVer As String

    sVer = HCGUI_binver(m_Parent.Binary)
    
    If sVer <> "" And sVer <> "Error" Then
        Hashcat_Version = "hashcat " & sVer & vbCrLf
    End If
    
End Function


Private Sub closeCommand_Click()

    Unload Me

End Sub
Public Function ShowInfo(fParent As fMain, Optional iMode As Long = 0, Optional sString As String = "", Optional sTitle As String = "")
Dim w As Long

    Set m_Parent = fParent
    
    'header picture
    Me.infoImage.Picture = fParent.headerPic.Picture
    Form_Resize
    
    'mode
    Select Case iMode
        Case 1: 'help
            Call zHelpBig("Commandline Help", m_Parent.Binary & " --help")
            
        Case 2: 'eula
            Call zHelpBig("Executing User License Agreement", m_Parent.Binary & " --eula")
        
        Case 3: 'text
            Call zHelpBig(sTitle, sString, 1)
            
        Case Else: 'about
            'caption
            infoCaption.Caption = ""
            'about text
            Me.lblVer.Caption = Hashcat_Version() & "hashcat-gui " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
            lblInfo.Caption = "hashcat © 2009 by atom," & vbCrLf & "hashcat-gui © 2009 by hakre " & vbCrLf & " all rights reserved." & vbCrLf & vbCrLf & "Software has been created for scientific, analyzation, demonstration and sportive reasons. It is a dual-use tool under federal german law in the meaning of the Convention on Cybercrime, Budapest, 23.XI.2001. Usage restricted to legal use."
                                            
    End Select
    
    Me.Show 1, fParent
    
    Unload Me


    


End Function

Private Sub Form_Activate()
    If Me.helpText.Visible Then
        Me.helpText.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 2 And KeyCode = 65 And Me.helpText.Visible Then
        'ctrl + a
        Me.helpText.SelStart = 0
        Me.helpText.SelLength = Len(Me.helpText.Text)
    End If
    
    If Shift = 2 And KeyCode = 67 And Me.helpText.Visible Then
        'ctrl + c
        Clipboard.Clear
        If Me.helpText.SelLength = 0 Then
            Call Clipboard.SetText(Me.helpText.Text)
        Else
            Call Clipboard.SetText(Me.helpText.SelText)
        End If
        
    End If

End Sub

Private Sub Form_Resize()
Dim t As Long
    
    'header size
    On Error Resume Next
        infoImageFrame.Move -6 * Screen.TwipsPerPixelX, -3 * Screen.TwipsPerPixelY, ScaleWidth + 12 * Screen.TwipsPerPixelX, infoImage.Height + 3 * Screen.TwipsPerPixelY
        infoImage.Left = -infoImageFrame.Left
    On Error GoTo 0
        
    'close frame
    closeFrame.Move 0, Me.ScaleHeight - closeFrame.Height, ScaleWidth
    closeCommand.Left = closeFrame.Width - closeCommand.Width - 8 * Screen.TwipsPerPixelX
    
    'info label position
    infoCaption.Left = ScaleWidth - infoCaption.Width - 16 * Screen.TwipsPerPixelX
    
    'text position
    On Error Resume Next
        t = infoImageFrame.Top + infoImageFrame.Height
        helpText.Move 0, t, ScaleWidth, closeFrame.Top - t
    On Error GoTo 0
    
    'wine logo position
    wineImage.Move Me.ScaleWidth - wineImage.Width - 8 * Screen.TwipsPerPixelX, Me.infoImageFrame.Top + Me.infoImageFrame.Height + 8 * Screen.TwipsPerPixelY
    
    

End Sub


Private Sub Form_Terminate()

    Set m_Parent = Nothing
    
End Sub

Private Sub wineImage_Click()
Dim r As Long
    
    r = ShellExecute(Me.hWnd, "open", "http://www.winehq.org/", "", "", SW_SHOWNORMAL)
    
End Sub

Private Sub wineImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub wineImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    SetCursor LoadCursor(0, IDC_HAND)
    
End Sub

Private Sub zHelpBig(sCaption As String, sCommand As String, Optional iType As Long = 0)
Dim w As Long

    'caption
    infoCaption.Caption = sCaption
    Me.Caption = "hashcat " & infoCaption.Caption
    
    'type: 0: command, 1: text
    If iType = 0 Then
        sCommand = GetCommandOutput(sCommand, , True)
    End If
    
    'help text
    With helpText
        .Text = sCommand
        If .Text = "" Then
            .Text = vbCrLf & "There was a problem launching the binary:" & vbCrLf & m_Parent.Binary
        End If
        .Visible = True
    End With
    
    'bigger window
    w = Me.ScaleWidth * 3
    If w > infoImage.Width Then w = infoImage.Width
    Me.Move Me.Left, Me.Top, w + Me.Width - Me.ScaleWidth, Me.ScaleHeight * 1.5 + Me.Height - Me.ScaleHeight
    

End Sub

