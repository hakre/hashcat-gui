VERSION 5.00
Begin VB.Form fRun 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Run"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton runCmd 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox runText 
      Height          =   1395
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Text            =   "fRun.frx":0000
      Top             =   180
      Width           =   4455
   End
End
Attribute VB_Name = "fRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

End Sub
Private Sub Form_Load()
    Me.Icon = HCGUI_Mainform.Icon
End Sub

Private Sub runCmd_Click()
Dim sCommand As String
Dim r As Long

    sCommand = Me.runText.Text
    
    On Error Resume Next
    r = Shell(sCommand, vbNormalFocus)
    
    If Err.Number <> 0 Then
        Me.Caption = "Error: " + CStr(Err.Number) + " " + Err.Description
    Else
        Me.Caption = "Run: " + CStr(r)
    End If
    
    
    On Error GoTo 0
    
    

End Sub
