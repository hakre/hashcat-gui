VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl uFileCommands 
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1755
   ScaleWidth      =   2250
   Begin MSComctlLib.Toolbar toolbar 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   582
      ButtonWidth     =   609
      Style           =   1
      ImageList       =   "images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_browse"
            Object.ToolTipText     =   "browse for file"
            ImageKey        =   "folder"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_edit"
            Object.ToolTipText     =   "edit file"
            ImageKey        =   "edit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList images 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uFileCommands.ctx":0000
            Key             =   "tripoint"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uFileCommands.ctx":015A
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "uFileCommands.ctx":06F4
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuRecentParent 
      Caption         =   "menuRecentParent"
      Begin VB.Menu menuRecent 
         Caption         =   "menuRecent"
         Index           =   0
      End
   End
End
Attribute VB_Name = "uFileCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' api defines
'

'autocomplete: http://vbnet.mvps.org/code/textapi/shautocomplete.htm
Private Declare Function SHAutoComplete Lib "shlwapi" _
  (ByVal hwndEdit As Long, _
   ByVal dwFlags As Long) As Long
   
'Includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
Private Const SHACF_FILESYSTEM  As Long = &H1

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
   (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, _
   ByVal lpszWindow As String) As Long

'
' vb stuff
'

Public Event Changed()
Public Event Click(sKey As String)

Private m_Control As Object
Private m_Enabled As Boolean
Private m_OldText As String
Private WithEvents m_Recent As cRecent
Attribute m_Recent.VB_VarHelpID = -1
Public Property Get Bar() As toolbar
    Set Bar = toolbar
End Property
Public Property Get BoundControl() As Control
    Set BoundControl = m_Control
End Property
Public Property Set BoundControl(Textbox As Control)
Dim hWnd As Long

    Set m_Control = Textbox
    
    Select Case TypeName(Textbox)
        Case "TextBox":
            Call SHAutoComplete(m_Control.hWnd, SHACF_FILESYSTEM)
        Case "ComboBox":
            hWnd = FindWindowEx(m_Control.hWnd, 0, "EDIT", vbNullString)
            Call SHAutoComplete(hWnd, SHACF_FILESYSTEM)
        Case Else:
    End Select

End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(bEnabled As Boolean)
    m_Enabled = bEnabled
    toolbar.Enabled = m_Enabled
End Property
Public Property Get HasRecent() As Boolean
    HasRecent = Not CBool(m_Recent Is Nothing)
End Property
Public Property Let HasRecent(bValue As Boolean)
    If bValue Then
        If m_Recent Is Nothing Then
            Set m_Recent = New cRecent
        End If
    Else
        Set m_Recent = Nothing
    End If
End Property
Public Sub Trigger(TheEvent As eAcTriggerEvents)
Dim Ctl As ComboBox
Dim t As String
Dim i As Long

    If TheEvent = ClickEvent Then
        If m_OldText <> m_Control.Text Then
            RaiseEvent Changed
            m_OldText = m_Control.Text
        End If
    End If

    If TheEvent = ChangeEvent Then
        If Not m_Control Is Nothing Then
            If m_OldText <> m_Control.Text Then
                RaiseEvent Changed
                m_OldText = m_Control.Text
            End If
        End If
    End If
End Sub
Public Property Get Recent() As cRecent
    Set Recent = m_Recent
End Property
Public Sub RecentTouch()
    If Not m_Recent Is Nothing Then
        If Len(m_Control.Text) Then
            m_Recent.Touch m_Control.Text
        End If
    End If
End Sub




Private Sub barcmd_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub m_Recent_Update()
Dim oFi As cFileinfo
Dim Ctl As ComboBox
Dim i As Long, m As Long
Dim t As String

    If TypeName(m_Control) = "ComboBox" Then
        'Fill the combobox with items
        Set Ctl = m_Control
        'Ctl.Clear
        m = Ctl.ListCount
        t = Ctl.Text
        If m > 0 Then
            For i = m To 1 Step -1
                If t = Ctl.list(i - 1) Then
                    Call Ctl.RemoveItem(i - 1)
                    Ctl.Text = t
                Else
                    Call Ctl.RemoveItem(i - 1)
                End If
            Next i
        End If
        For Each oFi In m_Recent
            Ctl.AddItem oFi.Path
        Next
        Ctl.Text = t
        Ctl.ListIndex = 0
    End If
    
End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    RaiseEvent Click(Button.Key)
End Sub

Private Sub UserControl_Initialize()

    'Enabled
    m_Enabled = True

End Sub
Private Sub UserControl_InitProperties()

    'Enabled
    Me.Enabled = True
    
    'Recent
    Me.HasRecent = True

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'Enabled
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    
    'Recent
    Me.HasRecent = PropBag.ReadProperty("HasRecent", True)
    
    
End Sub
Private Sub UserControl_Resize()
Dim w As Long, h As Long

    With toolbar
        w = .Left + .Width
        h = .Height
    End With
            
    On Error Resume Next
        Width = w
        Height = h
        
End Sub

Private Sub UserControl_Terminate()
    Set m_Recent = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Enabled
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    
    'Recent
    Call PropBag.WriteProperty("HasRecent", Me.HasRecent, True)

End Sub
