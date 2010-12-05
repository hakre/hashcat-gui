VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fSettings 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Select hashcat binary"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "fSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmd 
      Caption         =   "&Reset"
      Height          =   375
      Index           =   2
      Left            =   5220
      TabIndex        =   6
      Top             =   840
      Width           =   1275
   End
   Begin VB.CheckBox osCheck 
      Caption         =   "&WINE"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "compability layer for WINE regarding hashcat executeable"
      Top             =   900
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3900
      TabIndex        =   5
      Top             =   840
      Width           =   1275
   End
   Begin VB.ComboBox binText 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   5955
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin MSComctlLib.Toolbar binTextTB 
      Height          =   330
      Left            =   6060
      TabIndex        =   1
      Top             =   120
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "browse"
            Object.ToolTipText     =   "Browse for hashcat CLI binary"
            ImageKey        =   "tripoint"
         EndProperty
      EndProperty
   End
   Begin VB.Label binLabel 
      Caption         =   "Label1"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   4515
   End
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Parent As fMain
Public WithEvents BinFile As cFbHelper
Attribute BinFile.VB_VarHelpID = -1
Public Function Ask(fParent As Form)
    Set m_Parent = fParent
    Me.binText.Text = m_Parent.Binary
    Me.BinOs = m_Parent.BinaryOs
    Me.Show 1, m_Parent
End Function
Public Property Get BinOs() As eAcBinOs
    If osCheck.Value = vbChecked Then
        BinOs = Wine
    Else
        BinOs = Windows
    End If
End Property
Public Property Let BinOs(Os As eAcBinOs)
    If Os = Wine Then
        osCheck.Value = vbChecked
    Else
        osCheck.Value = vbUnchecked
    End If
End Property

Private Sub zUpdateBinLabel()
Dim sCaption As String
Dim cFile As New cFileinfo

    cFile.Path = binText.Text
    binLabel.ToolTipText = cFile.Path
    
    sCaption = cFile.Basename & ": "
    
    If cFile.Exists Then
        sCaption = sCaption & HCGUI_binver(cFile.Path)
    Else
        sCaption = sCaption & "File not found."
    End If
    
    If binLabel.Caption <> sCaption Then
        binLabel.Caption = sCaption
    End If

End Sub

Private Sub BinFile_Changed()

    Call zUpdateBinLabel

End Sub

Private Sub BinFile_Click(sKey As String)
    Select Case sKey
        Case "browse":
            Call zCmdBrowseBinary
    End Select
End Sub


Private Sub binText_Change()

    Call BinFile.Trigger(ChangeEvent)

End Sub

Private Sub binText_Click()

    Call BinFile.Trigger(ClickEvent)
    

End Sub

Private Sub binText_GotFocus()

    Call zUpdateBinLabel
    
End Sub

Private Sub binText_KeyDown(KeyCode As Integer, Shift As Integer)

    zUpdateBinLabel

End Sub

Private Sub binText_LostFocus()

    BinFile.RecentTouch

End Sub

Private Sub binTextTB_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call BinFile.TriggerClick(Button.Key)
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
        BinFile.RecentTouch
        If HCGUI_BinFile <> Me.binText.Text Then
            HCGUI_BinFile = Me.binText.Text
            Call INIWrite(HCGUI_BinFile, "default", "hashcat", HCGUI_Inifile)
        End If
        m_Parent.Binary = HCGUI_BinFile
        If m_Parent.BinaryOs <> Me.BinOs Then
            HCGUI_BinOs = Me.BinOs
            Call INIWrite(HCGUI_BinOs, "default", "wineenabled", HCGUI_Inifile)
            m_Parent.BinaryOs = HCGUI_BinOs
        End If
    End If
    
    If Index < 2 Then
        Unload Me
    End If
    
    If Index = 2 Then
        Me.binText.Text = HCGUI_bin_default()
    End If
End Sub

Private Sub Form_Load()

' bind binText to recent and autocomplete
Set BinFile = New cFbHelper
Me.binTextTB.ImageList = m_Parent.plainsImages
Me.binTextTB.Buttons("browse").Image = "folder-open"
Call BinFile.Init(Me.binText, Me.binTextTB, True)

' load recent from ini
Call HCGUI_recent_from_ini(BinFile.Recent, HCGUI_Inifile, "hashcatbinaries")



End Sub

Private Sub Form_Unload(Cancel As Integer)

' save recent to ini
Call HCGUI_recent_to_ini(BinFile.Recent, HCGUI_Inifile, "hashcatbinaries")

End Sub


Private Sub zCmdBrowseBinary()
Dim sFile As String

    sFile = Me.binText.Text

    sFile = HCGUI_bin_askfor(Me.hWnd, Me.BinOs, sFile)
    If Len(sFile) Then
        Me.binText.Text = sFile
        BinFile.RecentTouch
    End If
    
End Sub

