VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPlains 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Plainfile(s)"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6540
   Icon            =   "fPlainsTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar toolbar 
      Align           =   3  'Links ausrichten
      Height          =   4335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   7646
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "arrup"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "arrdown"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "folder"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageKey        =   "delete"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton closeCmd 
      Caption         =   "Close"
      Height          =   435
      Left            =   2220
      TabIndex        =   1
      Top             =   3600
      Width           =   1395
   End
   Begin MSComctlLib.ImageList images 
      Left            =   480
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":000C
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":0166
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":02C0
            Key             =   "broken"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":041A
            Key             =   "arrdown"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":0574
            Key             =   "arrup"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fPlainsTool.frx":06CE
            Key             =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView list 
      Height          =   2955
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5212
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Name            =   "Bitstream Vera Sans"
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
         Object.Width           =   5010
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "fPlains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    list.Move toolbar.Width, 0, Me.ScaleWidth - toolbar.Width, Me.ScaleHeight


End Sub

Private Sub list_BeforeLabelEdit(Cancel As Integer)

End Sub
