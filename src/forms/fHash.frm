VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fHash 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Hash Browser"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fHash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView resultList 
      Height          =   4335
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Software"
         Object.Width           =   4834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hash"
         Object.Width           =   3670
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mode"
         Object.Width           =   1129
      EndProperty
   End
   Begin VB.TextBox searchText 
      ForeColor       =   &H80000003&
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Search"
      Top             =   60
      Width           =   5535
   End
   Begin VB.CommandButton cmdClick 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Index           =   1
      Left            =   5100
      TabIndex        =   4
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton searchKill 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5700
      MaskColor       =   &H8000000F&
      Picture         =   "fHash.frx":014A
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "reset search"
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "fHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Data As drhashCollection
Private m_Parent As fMain
Private Function List_Search(sText As String)

    sText = Trim(sText)
    
    Me.resultList.ListItems.Clear
    
    If Len(sText) Then
        Call List_Fill(m_Data, True, sText)
    Else
        Call List_Fill(m_Data)
    End If

End Function

Public Property Set Parent(formParent As fMain)
    Set m_Parent = formParent
End Property
Public Property Get Parent() As fMain
    Set Parent = m_Parent
End Property

Private Sub cmdClick_Click(Index As Integer)
Dim iMode As Long
Dim sMode As String

    Select Case Index
        Case 0:
            If Me.resultList.SelectedItem Is Nothing Then
                cmdClick(0).Enabled = False
                Exit Sub
            End If
            
            sMode = Me.resultList.SelectedItem.ListSubItems.Item(2).Text
            If sMode = "%" Then
                ' MsgBox ""
                Exit Sub
            End If
            iMode = Val(sMode)
            Me.Parent.HashMode = iMode
    End Select
    
    Me.Parent.DrhashWin.Toggle
End Sub

Private Sub Form_Deactivate()

    If Me.Visible Then
        Me.Parent.DrhashWin.Toggle
    End If
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 117 Then
        If Me.Visible Then
            Me.Parent.DrhashWin.Toggle
        End If
    End If
End Sub


Private Sub Form_Load()

    Set m_Data = New drhashCollection
    m_Data.Csv = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    Call List_Fill(m_Data)
    
End Sub
Private Function List_Fill(oData As drhashCollection, Optional bSearch As Boolean = False, Optional sSearch As String = "")
Dim cList As ListView
Dim oHash As drhashEntry
Dim oLi As ListItem
Dim i As Long
Dim oEntries As Object

    Set cList = Me.resultList
    
    If Not bSearch Then
        Set oEntries = oData
    Else
        Set oEntries = oData.SearchTitle(sSearch)
    End If
    
    For Each oHash In oEntries
        i = i + 1
        Set oLi = cList.ListItems.Add(, "k" + CStr(i), oHash.Title)
        Call oLi.ListSubItems.Add(, , oHash.Col(1))
        Call oLi.ListSubItems.Add(, , oHash.Col(2))
    Next
    
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then 'x pressed, let's toggle
        Call cmdClick_Click(1)
        Cancel = True
    End If
End Sub
Private Sub resultList_DblClick()
    If cmdClick(0).Enabled Then
        cmdClick_Click (0)
    End If
End Sub
Private Sub resultList_GotFocus()
    If Not resultList.SelectedItem Is Nothing Then
        cmdClick(0).Enabled = True
    End If
End Sub
Private Sub resultList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdClick(0).Enabled = True
End Sub
Private Sub searchKill_Click()
    searchText.Text = ""
    searchText_Change
    searchText_LostFocus
    resultList.SetFocus
End Sub
Private Sub searchText_Change()
    'skip search on empty string / inactive textbox
    If Me.searchText.ForeColor = &H80000003 Then
        searchKill.Enabled = False
        Exit Sub
    End If
    'search the list
    searchKill.Enabled = True
    Call List_Search(searchText.Text)
End Sub
Private Sub searchText_GotFocus()
    If Me.searchText.ForeColor = &H80000003 Then
        'forecolor  &H80000003&
        Me.searchText.Text = ""
        Me.searchText.ForeColor = &H80000008
    End If
End Sub
Private Sub searchText_LostFocus()
'set to initial
If Me.searchText.ForeColor = &H80000008 And Me.searchText.Text = "" Then
    Me.searchText.ForeColor = &H80000003
    Me.searchText.Text = "Search"
End If
End Sub
