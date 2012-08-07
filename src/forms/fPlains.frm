VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPlains 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Wordlist Manager"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   15
   ClientWidth     =   6540
   Icon            =   "fPlains.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList images 
      Left            =   60
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":000C
            Key             =   "file2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":0166
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":0700
            Key             =   "folder2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":085A
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":0DF4
            Key             =   "folder-open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":138E
            Key             =   "broken"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":14E8
            Key             =   "arrdown2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":1642
            Key             =   "arrdown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":1BDC
            Key             =   "arrup2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":1D36
            Key             =   "arrup"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":22D0
            Key             =   "delete2"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":242A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":29C4
            Key             =   "pinout"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":2B1E
            Key             =   "pinin"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":2C78
            Key             =   "chkall"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlains.frx":2DD2
            Key             =   "chknone"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolbar 
      Height          =   3300
      Left            =   5700
      TabIndex        =   0
      Top             =   0
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   5821
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stick"
            Object.ToolTipText     =   "Pin window to prevent the automatical close."
            ImageKey        =   "pinout"
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_up"
            Object.ToolTipText     =   "Move entry up"
            ImageKey        =   "arrup"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_down"
            Object.ToolTipText     =   "Move entry down"
            ImageKey        =   "arrdown"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "check_all"
            Object.ToolTipText     =   "Check all"
            ImageKey        =   "chkall"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "check_none"
            Object.ToolTipText     =   "Uncheck all"
            ImageKey        =   "chknone"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_add"
            Object.ToolTipText     =   "Add wordlist files"
            ImageKey        =   "folder-open"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_del"
            Object.ToolTipText     =   "Remove selected entries from the list"
            ImageKey        =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView list 
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   5212
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "images"
      SmallIcons      =   "images"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Wordlist"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   8538
      EndProperty
   End
   Begin VB.Menu wordlistMenuParent 
      Caption         =   "&Wordlist"
      Visible         =   0   'False
      Begin VB.Menu wordlistMenu 
         Caption         =   "&Add wordlist..."
         Index           =   0
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "&Remove selected wordlists"
         Index           =   2
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Remove &nonexistant"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Clear list"
         Index           =   4
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Check &All"
         Index           =   6
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "&Uncheck All"
         Index           =   7
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "&Inverse Checkmarks"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Move to Top"
         Index           =   10
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Move to Bottom"
         Index           =   11
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Move Up"
         Index           =   12
      End
      Begin VB.Menu wordlistMenu 
         Caption         =   "Move Down"
         Index           =   13
      End
   End
End
Attribute VB_Name = "fPlains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Parent As fMain
Private m_Pinned As Boolean
Private m_MsgBox As Boolean

Public ListHelper As cLvHelper

Public Sub InitToMe(formParent As fMain)
    Set Me.list.SmallIcons = formParent.plainsList.SmallIcons
    Set Me.list.Icons = formParent.plainsList.Icons
    
    Call formParent.yWordlistMenuinit(Me.wordlistMenu)
    
End Sub
Public Property Get Parent() As fMain
    Set Parent = m_Parent
End Property
Public Property Set Parent(formParent As fMain)
    Set m_Parent = formParent
End Property
Public Property Get Pinned() As Boolean
    Pinned = m_Pinned
End Property
Public Property Let Pinned(bValue As Boolean)
    m_Pinned = bValue
    Select Case bValue
        Case False
            Me.toolbar.Buttons.Item("stick").Value = tbrUnpressed
        Case True
            Me.toolbar.Buttons.Item("stick").Value = tbrPressed
    End Select
    
End Property
'
' to be called if something has changed
'
Private Sub zChanged()

    'escalate change on main form
    m_Parent.Job_Change
    
End Sub

'close this toolwindow
Private Sub zClose()
    If Me.Visible Then
        Me.Parent.PlainsWin.Toggle
    End If
End Sub

Private Function zListCmd(sCmd As String) As Long
Dim sNumbered As String
Dim iDir As Long
Dim r As Long

    zListCmd = 1 'command ok per default, that reduces code, fail is set later on

    Select Case sCmd
        Case "cb_copy", "cb_cut", "cb_paste":
            r = newLvPlains(list, ListHelper, m_Parent).cmd(sCmd)
            If r And sCmd = "cb_cut" Then
                r = zListCmd("item_del")
            End If
            zListCmd = r
            
        Case "check_all", "check_none":
            iDir = 1
            If sCmd = "check_none" Then iDir = 0
            If (list.ListItems.Count > ListHelper.CheckCount And iDir = 1) Or (ListHelper.CheckCount > 0 And iDir = 0) Then
                ListHelper.CheckAll iDir
                m_Parent.plainsListHelper.CheckAll iDir
                Call zChanged
            End If
            
        Case "check_invert":
            If ListHelper.list.ListItems.Count > 0 Then
                ListHelper.CheckInverse
                m_Parent.plainsListHelper.CheckInverse
                Call zChanged
            End If
        
        Case "item_up", "item_down":
            If ListHelper.SelCount > 0 Then
                If sCmd = "item_down" Then
                    iDir = 1
                End If
                sNumbered = ListHelper.NumberedSelected
                r = ListHelper.MoveNumbered(sNumbered, iDir)
                Call m_Parent.plainsListHelper.MoveNumbered(sNumbered, iDir)
                If r > 0 Then
                    Call zChanged
                End If
            End If
            
        Case "item_totop", "item_tobottom":
            If ListHelper.SelCount > 0 Then
                If sCmd = "item_tobottom" Then
                    iDir = 1
                End If
                sNumbered = ListHelper.NumberedSelected
                r = ListHelper.MoveNumberedMultiple(sNumbered, iDir, list.ListItems.Count - 1)
                Call m_Parent.plainsListHelper.MoveNumberedMultiple(sNumbered, iDir, list.ListItems.Count - 1)
                If r > 0 Then
                    Call zChanged
                End If
            End If
            
        Case "item_del":
            If ListHelper.SelCount > 0 Then
                sNumbered = ListHelper.NumberedSelected
                r = ListHelper.RemoveNumbered(sNumbered)
                Call m_Parent.plainsListHelper.RemoveNumbered(sNumbered)
                If r > 0 Then
                    Call zChanged
                End If
            End If
            
        Case "item_add":
            zListCmd = m_Parent.plainsList_AddFileDialog(Me.hWnd)
        
        Case "item_clear":
            If list.ListItems.Count > 0 Then
                m_MsgBox = True
                If MsgBox("Remove all wordlist file(s)?", vbQuestion + vbYesNo + vbDefaultButton1, "Clear wordlists") = vbYes Then
                    list.ListItems.Clear
                    Call m_Parent.plainsList.ListItems.Clear
                    Call zChanged
                End If
            End If
            
        Case "select_all":
            ListHelper.SelectAll
            
        Case "select_invert":
            ListHelper.SelectInverse
    
        Case "stick":
            Me.Pinned = CBool(toolbar.Buttons.Item("stick").Value)
            
        Case Else:
            Stop
            zListCmd = 0
            
    End Select

End Function
Private Sub Form_Deactivate()
    If Not Me.Pinned And Not m_MsgBox Then
        zClose
    End If
    If m_MsgBox Then
        m_MsgBox = False
    End If
End Sub
Private Sub Form_Initialize()
    Set ListHelper = New cLvHelper
    Set ListHelper.list = Me.list
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        If Me.Visible Then
            Me.Parent.PlainsWin.Toggle
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        zClose
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'x pressed, let's toggle
        Me.Parent.PlainsWin.Toggle
        Cancel = True
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    toolbar.Move Me.ScaleWidth - toolbar.Width
    list.Move 0, 0, Me.ScaleWidth - toolbar.Width, Me.ScaleHeight
End Sub
Private Sub list_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim keystat(0 To 255) As Byte ' receives key status information for all keys
Dim retval As Long ' return value of function
Dim sNumbered As String

    retval = GetKeyboardState(keystat(0)) ' In VB, the array is passed by referencing element #0.
    
    'if ctrl is pressed, check/de-check all selected
    If (keystat(&H11) And &H80) = &H80 Then
        sNumbered = ListHelper.NumberedSelected
        Call ListHelper.CheckNumbered(sNumbered, Item.Checked)
        Call m_Parent.plainsListHelper.CheckNumbered(sNumbered, Item.Checked)
    Else
        m_Parent.plainsList.ListItems.Item(Item.Index).Checked = Item.Checked
    End If
    Call zChanged
    
End Sub
Private Sub list_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r As Long

    If Shift = 0 Then
        Select Case KeyCode
            Case 45: r = zListCmd("item_add")
            Case 46: r = zListCmd("item_del")
            Case 93: PopupMenu Me.wordlistMenuParent
        End Select
    ElseIf Shift = 2 Then
        Select Case KeyCode
            Case 17: 'ctrl pressed
            Case 65: r = zListCmd("select_all")
            Case 67: r = zListCmd("cb_copy")
            Case 73: r = zListCmd("select_invert")
            Case 86: r = zListCmd("cb_paste")
            Case 88: r = zListCmd("cb_cut")
            'Case Else: Stop
        End Select
        ElseIf Shift = 3 Then
        Select Case KeyCode
            Case 73: r = zListCmd("check_invert")
            Case 77: r = zListCmd("check_all")
            Case 85: r = zListCmd("check_none")
        End Select
    ElseIf Shift = 4 Then
        Select Case KeyCode
            Case 35: r = zListCmd("item_tobottom")
            Case 36: r = zListCmd("item_totop")
            Case 38: r = zListCmd("item_up")
            Case 40: r = zListCmd("item_down")
        End Select
    End If
    

End Sub




Private Sub list_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu Me.wordlistMenuParent
    End If
End Sub
Private Sub list_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
    If oledd_is_files(Data) Then
        For i = 1 To Data.Files.Count
            Call m_Parent.plainsList_AddFile(Data.Files.Item(i))
        Next
    End If
End Sub
Private Sub list_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4
    Else
        Effect = vbDropEffectNone
    End If
End Sub
Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim r As Long
        
    r = zListCmd(Button.Key)
    
End Sub
Private Sub wordlistMenu_Click(Index As Integer)
Dim sCmd As String
Dim r As Long
    
    Select Case Index
        Case 0: sCmd = "item_add"
        Case 2: sCmd = "item_del"
        Case 3: sCmd = "item_removeinexistant"
        Case 4: sCmd = "item_clear"
        Case 6: sCmd = "check_all"
        Case 7: sCmd = "check_none"
        Case 8: sCmd = "check_invert"
        Case 10: sCmd = "item_totop"
        Case 11: sCmd = "item_tobottom"
        Case 12: sCmd = "item_up"
        Case 13: sCmd = "item_down"
    End Select
    
    If Len(sCmd) Then
        r = zListCmd(sCmd)
    End If
End Sub

