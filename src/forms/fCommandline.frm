VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fCommandline 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Commandline"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   635
      ButtonWidth     =   1402
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "Images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Win"
            Key             =   "os_win"
            ImageKey        =   "win"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Linux"
            Key             =   "os_linux"
            ImageKey        =   "linux"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wine"
            Key             =   "os_wine"
            ImageKey        =   "wine"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Key             =   "copy"
            ImageKey        =   "file"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "close"
            ImageKey        =   "close"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList Images 
      Left            =   1740
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCommandline.frx":0000
            Key             =   "win"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCommandline.frx":015A
            Key             =   "linux"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCommandline.frx":02B4
            Key             =   "wine"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCommandline.frx":040E
            Key             =   "file"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fCommandline.frx":09A8
            Key             =   "close"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textCommand 
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Text            =   "fCommandline.frx":0F42
      Top             =   960
      Width           =   5595
   End
End
Attribute VB_Name = "fCommandline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Parent As fMain
Private m_Display As Long 'commandlineDisplay



Public Property Get commandlineDisplay() As Long
    commandlineDisplay = m_Display
End Property
Public Property Let commandlineDisplay(Mode As Long)
    If Mode < 0 Then Exit Property
    If Mode > 2 Then Exit Property
    m_Display = Mode
    'Me.displayOption(Mode).Value = True
    toolbar.Buttons(Mode + 1).Value = tbrPressed
    
End Property


Public Property Get Parent() As fMain
    Set Parent = m_Parent
End Property

Public Property Set Parent(Parent As fMain)
    Set m_Parent = Parent
End Property
Private Function zCmd(sCmd As String) As Long

    Select Case sCmd
        Case "os_win", "os_linux", "os_wine": zCmd = zView(sCmd)
        Case "copy": zCmd = zCopy()
        Case "close": zCmd = zWinHide()
        Case Else:
            zCmd = 0
    End Select

End Function

Private Function zCopy() As Long

    If Len(Me.textCommand.Text) Then
        Clipboard.Clear
        Call Clipboard.SetText(Me.textCommand.Text)
        zCopy = 1
    End If
    
End Function
Private Function zView(sCmd As String) As Long
Dim iIndex As Long

    iIndex = COL_index(sCmd, "os_win", "os_linux", "os_wine") - 1
    If iIndex > -1 And iIndex <> m_Display Then
        Me.Parent.commandlineDisplay = iIndex
        zView = 1
    End If

End Function

Private Function zWinHide() As Long
        If Me.Visible Then
            Me.Parent.CommandlineWin.Toggle
            Me.Parent.commandlineMenu(1).Checked = Me.Parent.CommandlineWin.Visible
            Me.Parent.menuView(2).Checked = Me.Parent.CommandlineWin.Visible
            zWinHide = 1
        End If
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 118 Or KeyCode = 27 Then
        Call zWinHide
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' X pressed
    If UnloadMode = 0 Then
        Cancel = True
        zWinHide
    End If
        
End Sub
Private Sub Form_Resize()
Dim minheight As Long

    minheight = toolbar.Height * 2 + Me.Height - Me.ScaleHeight

On Error Resume Next
    If Me.Height < minheight Then
        Me.Height = minheight
        Exit Sub
    End If
    
    textCommand.Move 0, Me.toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Me.toolbar.Height
    
End Sub



Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call zCmd(Button.Key)
End Sub


