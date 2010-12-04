VERSION 5.00
Begin VB.UserControl uFileBox 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   KeyPreview      =   -1  'True
   ScaleHeight     =   900
   ScaleWidth      =   4605
   Begin VB.Frame commandFrame 
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   855
      Begin VB.CommandButton command 
         Height          =   315
         Index           =   1
         Left            =   480
         Picture         =   "uFileBox.ctx":0000
         Style           =   1  'Grafisch
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton command 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.TextBox textText 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "uFileBox"
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

Public Property Get Text() As String
    Text = textText.Text
End Property
Public Property Let Text(Text As String)
    textText.Text = Text
End Property

Private Sub UserControl_Initialize()


    ' add autocomplete for some text-boxes
    Call SHAutoComplete(textText.hwnd, SHACF_FILESYSTEM)


End Sub

Private Sub UserControl_InitProperties()

    Me.Text = UserControl.Name


End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'Text
    Me.Text = PropBag.ReadProperty("Text", Name)
    
End Sub


Private Sub UserControl_Resize()
Dim w As Long

    w = ScaleWidth - commandFrame.Width
    
    On Error Resume Next
        textText.Width = w
        textText.Height = ScaleHeight
        commandFrame.Left = w
        
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Text
    Call PropBag.WriteProperty("Text", Me.Text, Name)

End Sub


