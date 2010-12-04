VERSION 5.00
Begin VB.Form fOutput 
   Caption         =   "Hashcat Output"
   ClientHeight    =   6060
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9600
   Icon            =   "fOutput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraSep 
      Height          =   75
      Left            =   -120
      TabIndex        =   1
      Top             =   300
      Width           =   99135
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "fOutput.frx":57E2
      Top             =   660
      Width           =   8475
   End
   Begin VB.Menu menuClear 
      Caption         =   "&Clear"
   End
End
Attribute VB_Name = "fOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarFilename As String


Private Sub Form_Resize()
Dim t As Long

On Error Resume Next
fraSep.Move -6 * Screen.TwipsPerPixelX, -3 * Screen.TwipsPerPixelY, Me.ScaleWidth + 12 * Screen.TwipsPerPixelX

t = fraSep.Top + fraSep.Height
txtOutput.Move 0, t, Me.ScaleWidth, Me.ScaleHeight - t

End Sub


Public Function ShowEx(filname As String, Optional Ownerform As Form = Null)

    Me.Filename = Filename
    Me.Caption = "hashcat " & Me.Filename
        
    Me.Show 0, Ownerform

End Function



Public Property Let Filename(Value As String)

     mvarFilename = Value

End Property

Public Property Get Filename() As String

    Filename = mvarFilename

End Property

