VERSION 5.00
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
      TabIndex        =   2
      ToolTipText     =   "compability layer for WINE regarding hashcat executeable"
      Top             =   900
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3900
      TabIndex        =   4
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton plainsCmd 
      Caption         =   "..."
      Height          =   315
      Left            =   6120
      MaskColor       =   &H80000000&
      TabIndex        =   1
      ToolTipText     =   "Browse"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox binText 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   5955
   End
   Begin VB.CommandButton cmd 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   3
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label binLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   5
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

Private Sub binText_Change()
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

Private Sub cmd_Click(Index As Integer)
    If Index = 0 Then
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
        
    End If
End Sub

Private Sub plainsCmd_Click()
Dim sFile As String

    sFile = Me.binText.Text

    sFile = HCGUI_bin_askfor(Me.hWnd, Me.BinOs, sFile)
    If Len(sFile) Then
        Me.binText.Text = sFile
    End If
    
End Sub
