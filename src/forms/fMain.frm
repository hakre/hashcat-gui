VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "hashcat-gui"
   ClientHeight    =   11070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame optFrame 
      Caption         =   "&Bruteforce Settings"
      Height          =   735
      Index           =   3
      Left            =   60
      TabIndex        =   29
      Top             =   7380
      Width           =   8415
      Begin VB.ComboBox bruteCharsText 
         Height          =   315
         ItemData        =   "fMain.frx":57E2
         Left            =   1200
         List            =   "fMain.frx":57EC
         OLEDropMode     =   1  'Manuell
         TabIndex        =   31
         Text            =   "abcdefghijklmnopqrstuvwxyz"
         ToolTipText     =   "charset for attack"
         Top             =   300
         Width           =   7095
      End
      Begin VB.Label zLbl 
         Caption         =   "Charset:"
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Saltfile"
      Height          =   735
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   4200
      Width           =   8415
      Begin VB.CheckBox saltFileCheck 
         Caption         =   "Saltfile:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "enable to use a saltfile"
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox saltFileText 
         Height          =   315
         ItemData        =   "fMain.frx":5832
         Left            =   1200
         List            =   "fMain.frx":5834
         OLEDropMode     =   1  'Manuell
         TabIndex        =   15
         ToolTipText     =   "charset for attack"
         Top             =   240
         Width           =   6375
      End
      Begin MSComctlLib.Toolbar saltFileTb 
         Height          =   330
         Left            =   7620
         TabIndex        =   16
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "plainsImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_browse"
               Object.ToolTipText     =   "browse for saltfile"
               ImageKey        =   "folder-open"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_edit"
               Object.ToolTipText     =   "edit saltfile"
               ImageKey        =   "edit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Toggle Settings"
      Height          =   735
      Index           =   4
      Left            =   60
      TabIndex        =   32
      Top             =   8220
      Width           =   8415
      Begin VB.TextBox toggleLenText 
         Alignment       =   2  'Zentriert
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "1"
         ToolTipText     =   "number of alphas in plain minimum"
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox toggleLenText 
         Alignment       =   2  'Zentriert
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   36
         Text            =   "16"
         ToolTipText     =   "number of alphas in plain maximum"
         Top             =   300
         Width           =   495
      End
      Begin VB.Label zLbl 
         Caption         =   "Number:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   33
         ToolTipText     =   "number of alphas in plain"
         Top             =   360
         Width           =   915
      End
      Begin VB.Label zLbl 
         Alignment       =   2  'Zentriert
         Caption         =   "-"
         Height          =   255
         Index           =   5
         Left            =   1740
         TabIndex        =   35
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame optFrame 
      Caption         =   "Password &Length"
      Height          =   735
      Index           =   2
      Left            =   60
      TabIndex        =   24
      Top             =   6540
      Width           =   8415
      Begin VB.TextBox bruteLenText 
         Alignment       =   2  'Zentriert
         Height          =   315
         Index           =   1
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "8"
         ToolTipText     =   "password length maximum"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox bruteLenText 
         Alignment       =   2  'Zentriert
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "1"
         ToolTipText     =   "password length minimum"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label zLbl 
         Alignment       =   2  'Zentriert
         Caption         =   "-"
         Height          =   255
         Index           =   3
         Left            =   1740
         TabIndex        =   27
         Top             =   300
         Width           =   255
      End
      Begin VB.Label zLbl 
         Caption         =   "Length:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   "password length"
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.PictureBox bottomFrame 
      Align           =   2  'Unten ausrichten
      BorderStyle     =   0  'Kein
      Height          =   1680
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   8550
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9090
      Width           =   8550
      Begin VB.ComboBox outFileText 
         Height          =   315
         Left            =   1260
         OLEDropMode     =   1  'Manuell
         TabIndex        =   41
         Top             =   120
         Width           =   6435
      End
      Begin VB.CheckBox outFileCheck 
         Caption         =   "&Outfile:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   180
         Width           =   1035
      End
      Begin VB.CheckBox viewerCheck 
         Caption         =   "Monitor Outfile"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "start outfile viewer with hashcat"
         Top             =   1320
         Value           =   1  'Aktiviert
         Width           =   2655
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Secret Button (&A)"
         Height          =   555
         Index           =   0
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Press this Button to start."
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame fraSep 
         Height          =   75
         Index           =   1
         Left            =   -120
         TabIndex        =   43
         Top             =   540
         Width           =   99135
      End
      Begin VB.Frame resourcesFrame 
         BorderStyle     =   0  'Kein
         Height          =   855
         Left            =   2880
         TabIndex        =   58
         Top             =   720
         Visible         =   0   'False
         Width           =   5625
         Begin VB.TextBox skipText 
            Alignment       =   2  'Zentriert
            Height          =   315
            Index           =   0
            Left            =   2580
            TabIndex        =   52
            Text            =   "0"
            ToolTipText     =   "skip number of words from wordfile summary"
            Top             =   60
            Width           =   2895
         End
         Begin VB.TextBox segmentText 
            Alignment       =   2  'Zentriert
            Height          =   315
            Left            =   900
            MaxLength       =   9
            TabIndex        =   49
            Text            =   "32"
            ToolTipText     =   "number of bytes (MB) to read from wordfiles at once"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox threadsText 
            Alignment       =   2  'Zentriert
            Height          =   315
            Left            =   900
            MaxLength       =   2
            TabIndex        =   47
            Text            =   "8"
            ToolTipText     =   "number of threads"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox skipText 
            Alignment       =   2  'Zentriert
            Height          =   315
            Index           =   1
            Left            =   2580
            TabIndex        =   54
            Text            =   "0"
            ToolTipText     =   "limit number of words from wordfile summary"
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label zLbl 
            Caption         =   "Skip:"
            Height          =   255
            Index           =   9
            Left            =   2040
            TabIndex        =   51
            Top             =   120
            Width           =   435
         End
         Begin VB.Label zLbl 
            Caption         =   "MB"
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   50
            Top             =   540
            Width           =   435
         End
         Begin VB.Label zLbl 
            Caption         =   "Segment:"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   48
            Top             =   540
            Width           =   915
         End
         Begin VB.Label zLbl 
            Caption         =   "Threads:"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Width           =   855
         End
         Begin VB.Label zLbl 
            Caption         =   "Limit:"
            Height          =   255
            Index           =   10
            Left            =   2040
            TabIndex        =   53
            Top             =   540
            Width           =   615
         End
      End
      Begin MSComctlLib.Toolbar outFileTb 
         Height          =   330
         Left            =   7740
         TabIndex        =   42
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "plainsImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_browse"
               Object.ToolTipText     =   "browse for outfile"
               ImageKey        =   "folder-open"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_edit"
               Object.ToolTipText     =   "outfile viewer"
               ImageKey        =   "fileb"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar hashModeToolbar 
      Height          =   330
      Left            =   8100
      TabIndex        =   12
      Top             =   3780
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "plainsImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "drhash"
            Object.ToolTipText     =   "Hash Browser"
            ImageKey        =   "tripoint"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList plainsImages 
      Left            =   120
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":5836
            Key             =   "file2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":5990
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":5F2A
            Key             =   "folder2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":6084
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":661E
            Key             =   "folder-open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":6BB8
            Key             =   "broken"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":6D12
            Key             =   "arrup"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":72AC
            Key             =   "arrdown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":7846
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":7DE0
            Key             =   "listsymbol"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":7F3A
            Key             =   "listdetail"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":8094
            Key             =   "tripoint2"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":81EE
            Key             =   "tripoint"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":8348
            Key             =   "dropdown"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":84A2
            Key             =   "drhash"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":8A3C
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":8FD6
            Key             =   "fileb"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":9570
            Key             =   "copy"
         EndProperty
      EndProperty
   End
   Begin VB.Frame optFrame 
      Caption         =   "&Rules (Hybrid Attack)"
      Height          =   1455
      Index           =   1
      Left            =   60
      TabIndex        =   17
      Top             =   4980
      Width           =   8415
      Begin VB.ComboBox ruleGenerateText 
         Height          =   315
         ItemData        =   "fMain.frx":9B0A
         Left            =   1860
         List            =   "fMain.frx":9B20
         TabIndex        =   22
         Text            =   "10 k"
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox ruleFileText 
         Height          =   315
         Left            =   1200
         OLEDropMode     =   1  'Manuell
         TabIndex        =   20
         Top             =   600
         Width           =   6435
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2295
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   300
         Width           =   2295
         Begin VB.OptionButton ruleOption 
            Caption         =   "Use &File:"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton ruleOption 
            Caption         =   "&Generate Rules:"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   21
            Top             =   720
            Width           =   1755
         End
         Begin VB.OptionButton ruleOption 
            Caption         =   "&None."
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1035
         End
      End
      Begin MSComctlLib.Toolbar ruleFileTB 
         Height          =   330
         Left            =   7620
         TabIndex        =   59
         Top             =   600
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "plainsImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_browse"
               Object.ToolTipText     =   "browse for rulesfile"
               ImageKey        =   "folder-open"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cmd_edit"
               Object.ToolTipText     =   "edit rulesfile"
               ImageKey        =   "edit"
            EndProperty
         EndProperty
      End
      Begin VB.Label zLbl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   23
         Top             =   1020
         Width           =   4575
      End
   End
   Begin VB.TextBox hashSeperator 
      Alignment       =   2  'Zentriert
      Height          =   315
      Left            =   3060
      MaxLength       =   1
      TabIndex        =   5
      Text            =   ":"
      ToolTipText     =   "Char that seperates salts from hashes in the hashfile"
      Top             =   1620
      Width           =   495
   End
   Begin VB.PictureBox headerPic 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      Picture         =   "fMain.frx":9B44
      ScaleHeight     =   1080
      ScaleWidth      =   9600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
   End
   Begin VB.Frame fraSep 
      Height          =   75
      Index           =   2
      Left            =   0
      TabIndex        =   56
      Top             =   900
      Width           =   99135
   End
   Begin VB.ComboBox hashModeCombo 
      Height          =   315
      ItemData        =   "fMain.frx":2B788
      Left            =   4980
      List            =   "fMain.frx":2B78A
      TabIndex        =   11
      Text            =   "hashModeCombo"
      ToolTipText     =   "Hashmode"
      Top             =   3780
      Width           =   3075
   End
   Begin VB.ComboBox recoveryModeCombo 
      Height          =   315
      ItemData        =   "fMain.frx":2B78C
      Left            =   1260
      List            =   "fMain.frx":2B78E
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      ToolTipText     =   "Recoverymode"
      Top             =   3780
      Width           =   2535
   End
   Begin VB.ComboBox hashFileText 
      Height          =   315
      Left            =   1260
      OLEDropMode     =   1  'Manuell
      TabIndex        =   2
      Text            =   "Test"
      Top             =   1200
      Width           =   6075
   End
   Begin VB.Frame fraSep 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   55
      Top             =   780
      Width           =   99135
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   37
      Top             =   10770
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   529
      SimpleText      =   "hashcat-gui"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14129
            Object.ToolTipText     =   "Commandline - dbl.-click to open"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   423
            MinWidth        =   423
            Object.ToolTipText     =   "hashcat version - dbl.-click to select"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   423
            MinWidth        =   423
            Object.ToolTipText     =   "GUI version - dbl.-click for info"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView plainsList 
      Height          =   1635
      Left            =   1260
      TabIndex        =   7
      Top             =   2040
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   2884
      View            =   2
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
      Icons           =   "plainsImages"
      SmallIcons      =   "plainsImages"
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
         Object.Width           =   3070
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   7832
      EndProperty
   End
   Begin MSComctlLib.Toolbar plainsToolbar 
      Height          =   330
      Left            =   6180
      TabIndex        =   8
      Top             =   1680
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "plainsImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_up"
            Object.ToolTipText     =   "Move wordlist up"
            ImageKey        =   "arrup"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_down"
            Object.ToolTipText     =   "Move wordlist down"
            ImageKey        =   "arrdown"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_add"
            Object.ToolTipText     =   "Add wordlist file"
            ImageKey        =   "folder-open"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "item_del"
            Object.ToolTipText     =   "Remove wordlist from list"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "list_mode"
            Object.ToolTipText     =   "Display mode (Columns / List)"
            ImageKey        =   "listsymbol"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "form"
            Object.ToolTipText     =   "Open wordlist manager"
            ImageKey        =   "tripoint"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar hashFileTb 
      Height          =   330
      Left            =   7440
      TabIndex        =   3
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "plainsImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_browse"
            Object.ToolTipText     =   "browse for hashfile"
            ImageKey        =   "folder-open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_edit"
            Object.ToolTipText     =   "edit hashfile"
            ImageKey        =   "edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmd_paste"
            Object.ToolTipText     =   "quick paste"
            ImageKey        =   "copy"
         EndProperty
      EndProperty
   End
   Begin VB.Label zLbl 
      Caption         =   "Hashlist Se&perator:"
      Height          =   255
      Index           =   14
      Left            =   1260
      TabIndex        =   4
      Top             =   1680
      Width           =   1755
   End
   Begin VB.Label zLbl 
      Caption         =   "Ha&sh:"
      Height          =   255
      Index           =   12
      Left            =   4080
      TabIndex        =   38
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label zLbl 
      Caption         =   "&Mode:"
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   9
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label zLbl 
      Caption         =   "&Hashfile:"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   1
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label zLbl 
      Caption         =   "&Wordlist(s):"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   975
   End
   Begin VB.Menu menuFileMain 
      Caption         =   "&File"
      Begin VB.Menu menuFileOp 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "&Recent Files"
         Index           =   2
         Begin VB.Menu menuFileRecent 
            Caption         =   "test"
            Index           =   0
         End
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "&Save"
         Index           =   4
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "Save &As..."
         Index           =   5
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu menuFileOp 
         Caption         =   "E&xit"
         Index           =   7
      End
      Begin VB.Menu menuFileDebug 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu menuFileDebug 
         Caption         =   "&Run..."
         Index           =   1
      End
      Begin VB.Menu menuFileDebug 
         Caption         =   "&Shell (cmd)..."
         Index           =   2
      End
      Begin VB.Menu menuFileDebug 
         Caption         =   "&Wineconsole..."
         Index           =   3
      End
   End
   Begin VB.Menu menuFileA 
      Caption         =   "&Settings"
      Begin VB.Menu menuFile 
         Caption         =   "Console stays open after execute [recommended]"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu menuFile 
         Caption         =   "&Show resource related settings"
         Index           =   1
      End
      Begin VB.Menu menuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu menuFile 
         Caption         =   "Select hashcat binary..."
         Index           =   3
      End
   End
   Begin VB.Menu commandlineMenuParent 
      Caption         =   "&Commandline"
      Begin VB.Menu commandlineMenu 
         Caption         =   "&Copy"
         Index           =   0
      End
      Begin VB.Menu commandlineMenu 
         Caption         =   "&Show"
         Index           =   1
      End
      Begin VB.Menu commandlineMenu 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu commandlineMenu 
         Caption         =   "&Display"
         Index           =   3
         Begin VB.Menu commandlineDisplayMenu 
            Caption         =   "Win"
            Index           =   0
         End
         Begin VB.Menu commandlineDisplayMenu 
            Caption         =   "Linux"
            Index           =   1
         End
         Begin VB.Menu commandlineDisplayMenu 
            Caption         =   "Wine"
            Index           =   2
         End
      End
   End
   Begin VB.Menu plainsMenu 
      Caption         =   "&Wordlist"
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "&Add wordlist..."
         Index           =   0
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "&Remove selected wordlists"
         Index           =   2
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Remove &nonexistant"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Clear list"
         Index           =   4
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Check &All"
         Index           =   6
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "&Uncheck All"
         Index           =   7
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "&Inverse Checkmarks"
         Index           =   8
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Move to Top"
         Index           =   10
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Move to Bottom"
         Index           =   11
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Move Up"
         Index           =   12
      End
      Begin VB.Menu plainsMenuEntry 
         Caption         =   "Move Down"
         Index           =   13
      End
   End
   Begin VB.Menu menuViewParent 
      Caption         =   "&View"
      Begin VB.Menu menuView 
         Caption         =   "Wordlist Manager"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu menuView 
         Caption         =   "Hash Browser"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu menuView 
         Caption         =   "Commandline"
         Index           =   2
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu helpMenu 
      Caption         =   "&Help"
      Begin VB.Menu helpMenuEntry 
         Caption         =   "&About..."
         Index           =   0
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "Commandline &Help..."
         Index           =   2
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "&EULA..."
         Index           =   3
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "Online Forum..."
         Index           =   4
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu helpMenuEntry 
         Caption         =   "Debug Info..."
         Index           =   6
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' api defines
'
Private Declare Sub InitCommonControls Lib "comctl32.dll" () 'xp manifest theme support

'
' application code
'
Private m_RuleMode As Long
Private m_Filename As String
Private m_Filedirty As Boolean
Private m_CommandlineLoaded As Boolean

Private m_Binary As String
Private m_BinaryOs As eAcBinOs
Private m_commandlineDisplay As Long

Public WithEvents RecentJobs As cRecent
Attribute RecentJobs.VB_VarHelpID = -1

Public CommandlineWin As cTwHelper
Public DrhashWin As cTwHelper
Public PlainsWin As cTwHelper

Private m_RecentMenu As New cRecentMenu
'Private RecentFiles As cRecentCollection
'Private m_menuRecent_Clicked As Long

Public plainsListHelper As cLvHelper

Public BruteChars As cFbHelper
Public WithEvents HashFile As cFbHelper
Attribute HashFile.VB_VarHelpID = -1
Public WithEvents OutFile As cFbHelper
Attribute OutFile.VB_VarHelpID = -1
Public WithEvents SaltFile As cFbHelper
Attribute SaltFile.VB_VarHelpID = -1
Public WithEvents RuleFile As cFbHelper
Attribute RuleFile.VB_VarHelpID = -1
Public Property Get Binary() As String
    Binary = m_Binary
End Property
Public Property Let Binary(Path As String)
Dim sText As String
Dim cPanel As Panel
Dim iVer As Long

    m_Binary = Path
        
    'update statusbar
    sText = HCGUI_binver(m_Binary)
    
    Set cPanel = statusBar.Panels.Item(2)
    cPanel.Text = sText & " " 'hashcat
    
    'update recovery modes combo
    If Val(sText) >= 0.32 Then
        iVer = 1
    End If
    
    Call zFillAttackmodes(Me.recoveryModeCombo, iVer)
    
    Commandline_Change

End Property
Public Property Let BinaryOs(Os As eAcBinOs)
Dim bFlag As Boolean
Dim i As Long

    m_BinaryOs = Os
    
    bFlag = CBool(Os = Wine)
    bFlag = False
    For i = 0 To 3
        menuFileDebug(i).Visible = bFlag
    Next i
    
    
    Commandline_Change
    
End Property

Public Property Get BinaryOs() As eAcBinOs
    BinaryOs = m_BinaryOs
End Property
Private Sub Commandline_Change()
Dim sCaption As String
Dim i As Long
Dim oLine As New cAcCommandlineBuilder

    'use commandline builder to build the commandline
    Set oLine.Form = Me
    sCaption = oLine.Commandline
    
    'update commandline display menu and toolwindow options
    Me.commandlineMenu(1).Checked = Me.CommandlineWin.Visible
    For i = 0 To 2
        Me.commandlineDisplayMenu(i).Checked = CBool(Me.commandlineDisplay = i)
    Next i
    
    'commandline
    Me.statusBar.Panels.Item(1).Text = sCaption
        
    'update toolwindow if loaded
    If Me.CommandlineWin.isLoaded Then
        Me.CommandlineWin.Toolwindow.Caption = "Commandline - " & Me.Executeable.Path
        Me.CommandlineWin.Toolwindow.commandlineDisplay = Me.commandlineDisplay
        Me.CommandlineWin.Toolwindow.textCommand = sCaption
    End If
        
End Sub

Public Property Let Dirty(bVal As Boolean)
Dim oFi As cFileinfo
Dim sCap As String

    'reflect setting
    m_Filedirty = bVal
    
    'change menu accordingly
    menuFileOp.Item(4).Enabled = bVal
    

    'reflect the changes for empty files
    If m_Filename = "" Then
    
        If bVal Then
            sCap = "* - "
        End If
        
        Me.Caption = sCap & "hashcat-gui"
                
        Exit Property
    End If
    
    
    'reflect the changes for existing files
    Set oFi = New cFileinfo
    oFi.Path = m_Filename
    
    sCap = oFi.Path
    If m_Filedirty Then
        sCap = sCap & "*"
    End If
    
    Me.Caption = sCap & " - hashcat-gui"
End Property

'
' getter for hashcat binary fileinfo
'
Public Property Get Executeable() As cFileinfo
Dim fileBinary As New cFileinfo
    fileBinary.Path = m_Binary
    Set Executeable = fileBinary
End Property
Public Function ExecuteHashcat()
Dim oBuilder As New cAcCommandlineBuilder
Dim sPrefix As String
Dim sSuffix As String
Dim r As Long
Dim oFi As New cFileinfo
Dim bStayOpen As Boolean
    
    'new style commandline building
    Set oBuilder.Form = Me
    
    'new window
    bStayOpen = CBool(Me.menuFile.Item(0).Checked)
    If bStayOpen Then
        sPrefix = "CMD /K """
        sSuffix = """"
    End If
    
    'wine compat layer
    If Me.BinaryOs = Wine Then
        sPrefix = "wineconsole " & sPrefix
        oBuilder.Display = 0 'windows
    Else
        oBuilder.Display = 2 'wine
    End If
    
    oBuilder.Prefix = sPrefix
    oBuilder.Suffix = sSuffix
    
    'change to current working directory
    oFi.Path = oBuilder.Binary
    On Error Resume Next
    Call ChDir(oFi.Dirname)
    Call ChDrive(oBuilder.Binary)
    On Error GoTo 0
    
    On Error Resume Next
    r = Shell(oBuilder.Commandline, vbNormalFocus)
    If Err.Number <> 0 Then
        MsgBox "An Error occured while executing hashcat:" & vbCrLf & vbCrLf & "#" & CStr(Err.Number) & " - " & Err.Description & vbCrLf & vbCrLf & oBuilder.Commandline, vbCritical, "Error"
    End If
    On Error GoTo 0
    
End Function

'
' ask for discarding dirty if applicable
'
Public Function FileAskDirty(Optional sReason As String = "Exit hashcat-gui") As Boolean
Dim iRet As Long

    FileAskDirty = True

    If m_Filedirty Then
        iRet = MsgBox("Discard changes?", vbYesNo Or vbDefaultButton2 Or vbQuestion, sReason)

        ' new file on yes
        If iRet = vbNo Then FileAskDirty = False
    End If
    
End Function
'
' ask for quit if applicable
'
Public Function FileAskQuit() As Boolean

    If m_Filedirty And m_Filename <> "" Then
        If Not FileAskDirty Then
            FileAskQuit = False
            Exit Function
        End If
    End If
            
    FileAskQuit = True
    
End Function

'
' load a job file
'
Public Function FileOpen(sFile As String)
    
    m_Filename = sFile
    Set Me.Job = HCGUI_job_from_ini(m_Filename)
    Me.RecentJobs.Touch m_Filename
    'set dirty flag
    Me.Dirty = False

End Function


'
Public Property Get HashMode() As Long
Dim iIndex As Long

    iIndex = hashModeCombo.ListIndex
    
    If iIndex < 0 Then
        HashMode = Val(hashModeCombo.Text)
        If HashMode < 0 Then HashMode = 0
    Else
        HashMode = hashModeCombo.ItemData(iIndex)
    End If

End Property
Public Property Let HashMode(iMode As Long)
Dim iIndex As Long

    iMode = MinMax(iMode, 0, 10000)
    
    iIndex = zHashmodeIndexInList(Me.hashModeCombo, iMode)
    
    If iIndex > 0 Then
        Me.hashModeCombo.ListIndex = iIndex - 1
        Exit Property
    End If
    
    
    'Call List_Fill(Me.hashModeCombo, iMode, )
    With Me.hashModeCombo
        Call .AddItem("Manual (" & CStr(iMode) & ")", 0)
        .ItemData(.NewIndex) = iMode
        .ListIndex = .NewIndex
    End With
    
End Property
Public Property Get HashModeHasSalt() As Boolean
    HashModeHasSalt = zHashmodeHasSalt(Me.HashMode)
End Property
' setter for this forms job
Public Property Set Job(oJob As cJob)
Dim iIndex As Long

    'attack
    iIndex = oJob.RecoveryMode
    If iIndex >= recoveryModeCombo.ListCount Then
        iIndex = recoveryModeCombo.ListCount - 1
    End If
    recoveryModeCombo.ListIndex = iIndex
    
    ' brute force options
    Me.bruteCharsText.Text = oJob.BruteChars
    Me.bruteLenText(0).Text = oJob.BruteLen.StringFrom
    Me.bruteLenText(1).Text = oJob.BruteLen.StringTo
    
    'hash
    hashFileText.Text = oJob.HashFile
    Me.HashMode = oJob.HashMode
    hashSeperator.Text = oJob.hashSeperator

    'out
    outFileText.Text = oJob.OutFile
    outFileCheck.Value = Abs(oJob.OutFile.Use)
    
    'plains
    plainsList.ListItems.Clear
    Call plainsList_AddPlains(oJob.Plains)
    If Me.PlainsWin.isLoaded Then
        Call Me.plainsListHelper.CopyList(PlainsWin.Toolwindow.list)
    End If
    
    'rules
    Me.ruleFileText.Text = oJob.RuleFile
    Me.ruleGenerateText.Text = oJob.RuleCount
    Me.RuleMode = oJob.RuleMode
    
    'saltfile
    saltFileText.Text = oJob.SaltFile
    saltFileCheck.Value = Abs(oJob.SaltFile.Use)
    
    'segment size
    segmentText.Text = CStr(oJob.Segment)
    
    'threads
    threadsText.Text = CStr(oJob.Threads)
    
    'toggle
    Me.toggleLenText(0).Text = oJob.ToggleLen.StringFrom
    Me.toggleLenText(1).Text = oJob.ToggleLen.StringTo
        
    'skip / limit
    skipText(0).Text = CStr(oJob.Skip)
    skipText(1).Text = CStr(oJob.Limit)

End Property

' getter for this forms job
Public Property Get Job() As cJob
Dim oJob As New cJob
    
    oJob.RuleFile = ruleFileText.Text
    oJob.RuleCount = Me.ruleGenerateText.Text
    oJob.RuleMode = Me.RuleMode
    
    oJob.HashFile = hashFileText.Text
    oJob.HashMode = Me.HashMode
    oJob.hashSeperator = Me.hashSeperator.Text
    
    ' recoverymode
    oJob.RecoveryMode = recoveryModeCombo.ListIndex
    
    ' brute force options
    oJob.BruteChars = Me.bruteCharsText.Text
    oJob.BruteLen.ValueFrom = Val(Me.bruteLenText(0).Text)
    oJob.BruteLen.ValueTo = Val(Me.bruteLenText(1).Text)
    
    Set oJob.Plains = Me.Plains
    
    ' outfile
    oJob.OutFile.Value = outFileText.Text
    oJob.OutFile.Use = outFileCheck.Value
    
    'saltfile
    oJob.SaltFile.Value = saltFileText.Text
    oJob.SaltFile.Use = saltFileCheck.Value
    
    'segment size
    oJob.Segment = Val(segmentText.Text)
    
    'threads
    oJob.Threads = Val(threadsText.Text)
    
    'toggle
    oJob.ToggleLen.ValueFrom = Val(Me.toggleLenText(0).Text)
    oJob.ToggleLen.ValueTo = Val(Me.toggleLenText(1).Text)
    
    'skip / limit
    oJob.Skip = Cnv_Str2Dec(skipText(0).Text)
    oJob.Limit = Cnv_Str2Dec(skipText(1).Text)
        
Set Job = oJob
End Property
' local helper function to reflect data change
Public Function Job_Change()
    Me.Dirty = True
    
    'reflect the changes visuall in new textbox
    'TODO becomes obsolete somehow in the future
    'Dim oPlains As cPlains
    'Set oPlains = HCGUI_Plains_Expand(Me.Plains)
    'plainsText.Text = oPlains.Caption
    Commandline_Change
    

End Function
Public Property Get commandlineDisplay() As Long
    commandlineDisplay = m_commandlineDisplay
End Property

Public Property Let commandlineDisplay(Mode As Long)
    m_commandlineDisplay = Mode
    Commandline_Change
    
End Property
Private Function param_nummin(sParams As String, sData As String, sFlag As String, iDefault As Long, Optional iMin As Long = 0)
Dim iValue As Long

    If Len(sData) > 9 Then
        sData = Left(sData, 9)
    End If

    iValue = Round(Val(sData), 0)
            
    If iValue > iMin And iValue <> iDefault Then
        sParams = sParams & " " & sFlag & " " & CStr(iValue)
    End If
        
    param_nummin = sParams
    
End Function
'plains getter
Public Property Get Plains() As cPlains
Dim oLi As ListItem
Dim oPlain As cPlainfile
Dim oPlains As New cPlains

    For Each oLi In Me.plainsList.ListItems
        Set oPlain = New cPlainfile
        oPlain.FileName = oLi.ToolTipText
        oPlain.Checked = oLi.Checked
        Call oPlains.Add(oPlain)
    Next
    
Set Plains = oPlains
End Property
Private Function ini_load(sFile As String)

    'recent
    Call HCGUI_recent_from_ini(Me.RecentJobs, sFile)
    
    'recent hashes
    Call HCGUI_recent_from_ini(HashFile.Recent, sFile, "hash")
    
    'recent rules
    Call HCGUI_recent_from_ini(RuleFile.Recent, sFile, "rule")
    
    'recent out
    Call HCGUI_recent_from_ini(OutFile.Recent, sFile, "out")
    
    'recent saltfiles
    Call HCGUI_recent_from_ini(SaltFile.Recent, sFile, "salts")
    
    'showres
    If INIReadInt("default", "showres", sFile) Then
        Call menuFile_Click(1)
    End If
    
    'nostayopen
    If INIReadInt("default", "nostayopen", sFile) Then
        Me.menuFile.Item(0).Checked = False
    End If
    
    'monitor outfile
    If INIReadInt("default", "nomonitoroutfile", sFile) Then
        Me.viewerCheck.Value = False
    End If
    
    'job
    Set Me.Job = HCGUI_job_from_ini(sFile)
    Me.Dirty = False
    Commandline_Change
    
    
End Function


Private Function ini_save(sFile As String)

    'job
    Call HCGUI_job_to_ini(Me.Job, sFile)
    
    'nostayopen
    Call INIWrite(1 - Abs(Me.menuFile.Item(0).Checked), "default", "nostayopen", sFile)
    
    'monitor outfile
    Call INIWrite(1 - Abs(Me.viewerCheck.Value), "default", "nomonitoroutfile", sFile)
    
    
    'showres
    Call INIWrite(Abs(Me.menuFile.Item(1).Checked), "default", "showres", sFile)
    
    'recent files
    Call HCGUI_recent_to_ini(Me.RecentJobs, sFile)
    
    'recent out
    Call HCGUI_recent_to_ini(OutFile.Recent, sFile, "out")
    
    'recent rules
    Call HCGUI_recent_to_ini(RuleFile.Recent, sFile, "rule")
    
    'recent hashes
    Call HCGUI_recent_to_ini(HashFile.Recent, sFile, "hash")
    
    'recent saltfiles
    Call HCGUI_recent_to_ini(SaltFile.Recent, sFile, "salts")
    
    
End Function

Public Function plainsList_AddFileDialog(Optional ByVal hWnd As Long = 0) As Long
Dim cc As New cCommonDialog
Dim sFile As String
Dim sDir As String
Dim sInitDir As String


    If hWnd = 0 Then
        hWnd = Me.hWnd
    End If
    
    sInitDir = HCGUI_directory(101) 'dicts directory
    
    sDir = "Add directory by leaving in this text"
    sFile = sDir
    
    If cc.VBGetOpenFileName(sFile, , False, True, , True, _
                "All Files (*.*)|*.*|Wordlists|*.dic;*.txt|Dictionary files (*.dic)|*.dic|Text files (*.txt)|*.txt|Plain files (*plain*)|*plain*", _
                , sInitDir, "Add wordlist", , hWnd, OFN_NOVALIDATE Or OFN_HideReadOnly Or OFN_NOTESTFILECREATE Or OFN_EXTENSIONDIFFERENT Or OFN_AllowMultiselect Or OFN_Explorer) Then
                
        If Right(sFile, Len(sDir) + 1) = "\" & sDir Then
            'add directory
            Call plainsList_AddFile(Left(sFile, Len(sFile) - Len(sDir) - 1))
        Else
            'add file
            Call plainsList_AddFiles(cc.Files)
            'Dim oPlain As cPlainfile
            'Dim oPlains As New cPlains
            'If cc.Files.Count > 0 Then
            '    For Each vFile In cc.Files
            '        Set oPlain = New cPlainfile
            '        oPlain.External = vFile
            '        oPlain.Checked = True
            '        Call oPlains.Add(oPlain)
            '        Set oPlain = Nothing
            '        'Call plainsList_AddFile(vFile)
            '    Next
            '    Call plainsList_AddPlains(oPlains)
            '    zPlainsChanged
            'End If
        End If
    End If

End Function

Public Sub plainsList_AddFile(ByVal sFilename As String)
Dim oPlain As New cPlainfile
Dim oPlains As New cPlains

    oPlain.External = sFilename
    oPlain.Checked = True
    Call oPlains.Add(oPlain)
    Call plainsList_AddPlains(oPlains)
    zPlainsChanged
    
End Sub

Public Sub plainsList_AddFiles(colFiles As Collection)
Dim oPlain As cPlainfile
Dim oPlains As New cPlains
Dim vFile As Variant

    If colFiles.Count > 0 Then
        For Each vFile In colFiles
            Set oPlain = New cPlainfile
            oPlain.External = vFile
            oPlain.Checked = True
            Call oPlains.Add(oPlain)
        Next
        Call plainsList_AddPlains(oPlains)
        zPlainsChanged
    End If

End Sub


Public Function plainsList_AddPlains(oPlains As cPlains) As Long
Dim oList As ListView
Dim oList2 As ListView
Dim bSecond As Boolean
Dim oLi As ListItem
Dim oLi2 As ListItem
Dim oPlain As cPlainfile
Dim sImage As String
Dim iCount As Long
Dim sKey As String
Dim oFi As cFileinfo
Dim iMouseOld As Long


    iMouseOld = Screen.MousePointer
    Screen.MousePointer = vbHourglass
   
    
    Set oList = Me.plainsList
    If Me.PlainsWin.isLoaded Then
        Set oList2 = Me.PlainsWin.Toolwindow.list
        bSecond = True
    End If
    
    'iteration
    For Each oPlain In oPlains
        Set oFi = oPlain.Fileinfo
        'take only existing files
        If oFi.Exists Then
                      
            sKey = "k" + oFi.FullPath
            
            'Set oLi = oList.FindItem(oPlain.Short)
            'If bSecond Then
            '    Set oLi2 = oList2.FindItem(oPlain.Short)
            'End If
            
            Set oLi = plainsList_ByKey(oList, sKey)
            If bSecond Then
                Set oLi2 = plainsList_ByKey(oList2, sKey)
            End If
            
            'if the item doesn't already exists, add it
            If oLi Is Nothing Then
            
                'image
                sImage = "broken"
                If oFi.isDir Then
                    sImage = "folder"
                ElseIf oFi.isFile Then
                    sImage = "file"
                End If
                
                'add item
                Set oLi = oList.ListItems.Add(, sKey, oPlain.Short, sImage, sImage)
                oLi.SubItems(1) = oFi.Path
                If bSecond Then 'add item to second list, data will be copied at the end of the for-loop
                    Set oLi2 = oList2.ListItems.Add()
                End If
                iCount = iCount + 1
            End If
            
            'check dropped entries based on input data
            oLi.Checked = oPlain.Checked
            oLi.ToolTipText = oPlain.FileName
            If bSecond Then
                Call Me.plainsListHelper.CopyItemData(oLi, oLi2)
            End If
        End If
    Next
    
    Screen.MousePointer = iMouseOld

plainsList_AddPlains = iCount
End Function
Public Function plainsList_ByKey(cList As ListView, sKey As String) As ListItem
Dim oItem As ListItem

    On Error Resume Next
        Set plainsList_ByKey = cList.ListItems.Item(sKey)
    On Error GoTo 0

End Function
Private Function List_Fill(cList As Control, iData As Long, sItem As String)
    Call cList.AddItem(sItem)
    cList.ItemData(cList.NewIndex) = iData
End Function


Public Function plainsList_Cmd(sCmd As String) As Long
    plainsList_Cmd = zPlainsListCmd(sCmd)
End Function

Public Property Get RuleGenerate() As String
Dim s As String

    s = Me.ruleGenerateText.Text
    
    s = LCase(s)
    s = Trim(s)
    
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")
    
    s = Replace(s, "m", "000k")
    s = Replace(s, "k", "000")
    
    s = CStr(Int(Val(s)))
         
RuleGenerate = s
End Property

Public Property Get RuleMode() As Long
Dim i As Long

    For i = 0 To 2
        If Me.ruleOption(i).Value Then
            RuleMode = i
            Exit Property
        End If
    Next
    
    RuleMode = 0 'defaults to none
    
End Property

Public Property Let RuleMode(iValue As Long)

    If iValue < 0 Then iValue = 0
    If iValue > 2 Then iValue = 2
    
    If m_RuleMode <> iValue Then
        m_RuleMode = iValue
        Job_Change
    End If
    
    Me.ruleOption(iValue).Value = True
    
End Property



'Add wordlist shortcut descriptions to such a menu (we got to, that's why public)
Friend Function yWordlistMenuinit(oMenuArray As Object)
    
    Call Menu_SetShortcut(oMenuArray(0), "Ins")
    Call Menu_SetShortcut(oMenuArray(2), "Del")
    
    Call Menu_SetShortcut(oMenuArray(6), "Ctrl+Shift+M")
    Call Menu_SetShortcut(oMenuArray(7), "Ctrl+Shift+U")
    Call Menu_SetShortcut(oMenuArray(8), "Ctrl+Shift+I")
    
    
    Call Menu_SetShortcut(oMenuArray(10), "Alt+Pos1")
    Call Menu_SetShortcut(oMenuArray(11), "Alt+End")
    Call Menu_SetShortcut(oMenuArray(12), "Alt+Up")
    Call Menu_SetShortcut(oMenuArray(13), "Alt+Down")

End Function



'
'create debug output
'
Private Function zDebugInfo() As String
Dim t As String
Dim sVer As String
Dim sCmd As String
Dim sFile As String
Dim sOutfileviewer As String
Dim oLine As New cAcCommandlineBuilder

    t = ""
    
    t = t & "hashcat-gui debuginfo v0.1" & vbCrLf
    t = t & "--------------------------" & vbCrLf & vbCrLf
    
    
    sVer = HCGUI_binver(Me.Binary)
    sOutfileviewer = HCGUI_directory(0) & "outfileviewer.exe"
    
    t = t & "hashcat ........ v" & sVer & vbCrLf
    t = t & "hashcat-gui .... v" & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision) & vbCrLf
    t = t & "outfileviewer .. v" & WINAPI_GetFileVersion(sOutfileviewer, 1) & vbCrLf
    t = t & vbCrLf
    
    Set oLine.Form = Me
    sCmd = oLine.Commandline
    t = t & "hashcat cmd:" & vbCrLf
    t = t & sCmd & vbCrLf & vbCrLf
    
    sFile = outFileText.Text
    sCmd = """" & sOutfileviewer & """" & " " & paramsafe(sFile, Windows)
    t = t & "outfileviewer cmd:" & vbCrLf
    t = t & sCmd & vbCrLf & vbCrLf

    
zDebugInfo = t
End Function

'
' fill attackmode combo
'
' can be set to a specific ver(sion):
'
'   0 : (default) 4 Attackmodes
'   1 : plus Permutation
Private Function zFillAttackmodes(ctlList As ComboBox, Optional ByVal ver As Long = 0)
Dim iIndex As Long

    iIndex = ctlList.ListIndex
    Call ctlList.Clear
    Call ctlList.AddItem("Straight-Words")
    Call ctlList.AddItem("Combination-Words")
    Call ctlList.AddItem("Toggle-Case")
    Call ctlList.AddItem("Brute-Force")
    
    If ver = 1 Then
        Call ctlList.AddItem("Permutation")
    End If
    
    If iIndex >= ctlList.ListCount Then
        iIndex = ctlList.ListCount - 1
    End If
    
    If iIndex > -1 Then
        ctlList.ListIndex = iIndex
    End If

End Function
' fill hashmode combo
Private Function zFillHashmodes(ctlList As ComboBox)

    Call List_Fill(ctlList, 0, "MD5")
    Call List_Fill(ctlList, 1, "md5($pass.$salt)")
    Call List_Fill(ctlList, 2, "md5($salt.$pass)")
    Call List_Fill(ctlList, 3, "md5(md5($pass))")
    Call List_Fill(ctlList, 4, "md5(md5(md5($pass)))")
    Call List_Fill(ctlList, 5, "md5(md5($pass).$salt)")
    Call List_Fill(ctlList, 6, "md5(md5($salt).$pass)")
    Call List_Fill(ctlList, 7, "md5($salt.md5($pass))")
    Call List_Fill(ctlList, 8, "md5($salt.$pass.$salt)")
    Call List_Fill(ctlList, 9, "md5(md5($salt).md5($pass))")
    Call List_Fill(ctlList, 10, "md5(md5($pass).md5($salt))")
    Call List_Fill(ctlList, 11, "md5($salt.md5($salt.$pass))")
    Call List_Fill(ctlList, 12, "md5($salt.md5($pass.$salt))")
    Call List_Fill(ctlList, 30, "md5($salt.chr(0).$pass)")
    Call List_Fill(ctlList, 31, "md5(md5(strtoupper(md5($pass)))")
    Call List_Fill(ctlList, 100, "SHA1")
    Call List_Fill(ctlList, 101, "sha1($pass.$salt)")
    Call List_Fill(ctlList, 102, "sha1($salt.$pass)")
    Call List_Fill(ctlList, 103, "sha1(sha1($pass))")
    Call List_Fill(ctlList, 104, "sha1(sha1(sha1($pass)))")
    Call List_Fill(ctlList, 105, "sha1(strtolower($username).$pass)")
    Call List_Fill(ctlList, 200, "MySQL")
    Call List_Fill(ctlList, 300, "MySQL4.1/MySQL5")
    Call List_Fill(ctlList, 400, "MD5(Wordpress)")
    Call List_Fill(ctlList, 400, "MD5(phpBB3)")
    Call List_Fill(ctlList, 400, "PHPASS")
    Call List_Fill(ctlList, 500, "MD5(Unix)")
    Call List_Fill(ctlList, 600, "SHA-1(Base64)")
    Call List_Fill(ctlList, 700, "SSHA-1(Base64)")
    Call List_Fill(ctlList, 800, "SHA-1(Django)")
    Call List_Fill(ctlList, 900, "MD4")
    Call List_Fill(ctlList, 1000, "NTLM")
    Call List_Fill(ctlList, 1100, "Domain Cached Credentials")
    Call List_Fill(ctlList, 1200, "MD5(Chap)")
    Call List_Fill(ctlList, 1300, "MSSQL")
    
    Dim oHashes As New drhashCollection
    Dim oHash As drhashEntry
    oHashes.Csv = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    For Each oHash In oHashes
        If Len(oHash.Title) And oHash.Col(2) <> "%" And oHash.Col(5) <> "algo" Then
            Call List_Fill(ctlList, CInt(Val(oHash.Col(2))), oHash.Title)
        End If
    Next
    
End Function
'Init the menus of this form
'used on form initialization
Private Function zFormMenuinit() As Long
    
    Call Menu_SetShortcut(menuFileOp(0), "Ctrl+N")
    Call Menu_SetShortcut(menuFileOp(1), "Ctrl+O")
    Call Menu_SetShortcut(menuFileOp(4), "Ctrl+S")
    Call Menu_SetShortcut(menuFileOp(5), "Ctrl+T")
    Call Menu_SetShortcut(menuFileOp(7), "Ctrl+W")
    
    Call yWordlistMenuinit(Me.plainsMenuEntry)
End Function


Private Function zHashmodeHasSalt(ByVal iMode As Long) As Boolean
Dim iIndex As Long
Dim sMode As String

    iIndex = zHashmodeIndexInList(Me.hashModeCombo, iMode)
    
    If iIndex > 0 Then
        sMode = Me.hashModeCombo.list(iIndex - 1)
        If InStr(sMode, "$salt") > 0 Then
            zHashmodeHasSalt = True
        End If
    End If
    
End Function

Private Function zHashmodeIndexInList(cList As Control, ByVal iMode As Long) As Long
Dim c As Long, i As Long

    c = cList.ListCount
        
    If c > 0 Then
        c = c - 1
        For i = 0 To c
            If cList.ItemData(i) = iMode Then
                zHashmodeIndexInList = i + 1
                Exit Function
            End If
        Next i
    End If

End Function


Private Sub zOptFrames(sEnabled As String)
Dim i As Long
Dim a As Long
Dim iFrames() As Long, vFrame As Variant
Dim bFlg As Boolean
Dim oFrame As Frame, oPrev As Frame
Dim iHeight As Long
Dim iPixels As Long

    iFrames = COL_arrayNumbered(sEnabled)
    
    iPixels = 8
    
    Set oPrev = Nothing
    
    For i = 0 To 4
        bFlg = CBool(COL_arrayIndex(i, iFrames))
        Set oFrame = optFrame.Item(i)
        
        'set visibility
        oFrame.Visible = bFlg
        
        'set position
        If CBool(bFlg) Then
            'need of move / something to orient on?
            If i > 0 Then
                If oPrev Is Nothing Then
                    oFrame.Move optFrame.Item(0).Left, optFrame.Item(0).Top
                Else
                    oFrame.Move oPrev.Left, oPrev.Top + oPrev.Height + iPixels * Screen.TwipsPerPixelY
                End If
            End If
            
            Set oPrev = oFrame
        End If
    Next i
    
    ' get top plus total height value
    If oPrev Is Nothing Then
        iHeight = optFrame.Item(0).Top
    Else
        iHeight = oPrev.Top + oPrev.Height
    End If
    
    'iHeight = iHeight + iPixels * Screen.TwipsPerPixelY
    
    ' set form size
    On Error Resume Next
        Me.Height = Me.Height - Me.ScaleHeight + iHeight + bottomFrame.Height + statusBar.Height
    On Error GoTo 0

End Sub

Private Sub zOptFramesChange()
Dim iRecoverymode As Long
Dim sFrames As String

    iRecoverymode = recoveryModeCombo.ListIndex
    
    Select Case iRecoverymode
        Case 0, 1:
            sFrames = "1"
            
        Case 2:
            sFrames = "1,4"
        
        Case 3:
            sFrames = "2,3"
            
        Case 4:
            sFrames = "2"
            
    End Select
    
    If Me.HashModeHasSalt = True Then
        sFrames = sFrames & ",0"
    End If
    
    Call zOptFrames(sFrames)

End Sub


Private Sub zPlainsDeleteCmd()
Dim sNumbered As String

    Stop
    'use zPlainsListCmd("item_del") instead

End Sub
Private Sub zPlainsEditorCmd()
Dim p1 As POINTAPI
Dim w As Single

    If Me.PlainsWin.isLoaded Then
        Me.PlainsWin.Toggle
    Else
        p1 = POS_control_bottomright(Me.plainsToolbar)
        w = Me.ScaleWidth - 2 * (Me.ScaleWidth - (plainsToolbar.Left + plainsToolbar.Width))
        PlainsWin.Toolwindow.Move p1.x * Screen.TwipsPerPixelX - w, p1.y * Screen.TwipsPerPixelY, w, Me.ScaleHeight - Me.plainsToolbar.Top - Me.plainsToolbar.Height - Me.statusBar.Height - 1 * (Me.ScaleWidth - (plainsToolbar.Left + plainsToolbar.Width))
        Call PlainsWin.Toolwindow.InitToMe(Me)
        Call Me.plainsListHelper.CopyList(PlainsWin.Toolwindow.list)
        Me.PlainsWin.Toggle
    End If
    
End Sub

Private Function zPlainsListCmd(sCmd As String) As Long
Dim sNumbered As String
Dim iDir As Long
Dim r As Long

    zPlainsListCmd = 1 'command ok per default, that reduces code, fail is set later on

    Select Case sCmd
        Case "cb_copy", "cb_cut", "cb_paste":
            r = newLvPlains(plainsList, plainsListHelper, Me).cmd(sCmd)
            If r And sCmd = "cb_cut" Then
                r = zPlainsListCmd("item_del")
            End If
            zPlainsListCmd = r
    
        Case "check_all", "check_none":
            iDir = 1
            If sCmd = "check_none" Then iDir = 0
            If (plainsList.ListItems.Count > plainsListHelper.CheckCount And iDir = 1) Or (plainsListHelper.CheckCount > 0 And iDir = 0) Then
                plainsListHelper.CheckAll iDir
                If PlainsWin.isLoaded Then
                    PlainsWin.Toolwindow.ListHelper.CheckAll iDir
                End If
                Call zPlainsChanged
            End If
        
        Case "check_invert":
            If plainsListHelper.list.ListItems.Count > 0 Then
                plainsListHelper.CheckInverse
                If PlainsWin.isLoaded Then
                    PlainsWin.Toolwindow.ListHelper.CheckInverse
                End If
                Call zPlainsChanged
            End If
            
        Case "form":
           Call zPlainsEditorCmd
           
        Case "item_up", "item_down": 'move selected items
            If plainsListHelper.SelCount > 0 Then
                If sCmd = "item_down" Then
                    iDir = 1
                End If
                sNumbered = plainsListHelper.NumberedSelected
                Call plainsListHelper.MoveNumbered(sNumbered, iDir)
                If PlainsWin.isLoaded Then
                    Call PlainsWin.Toolwindow.ListHelper.MoveNumbered(sNumbered, iDir)
                End If
                zPlainsChanged
            End If
            
        Case "item_totop", "item_tobottom":
            If plainsListHelper.SelCount > 0 Then
                If sCmd = "item_tobottom" Then
                    iDir = 1
                End If
                sNumbered = plainsListHelper.NumberedSelected
                r = plainsListHelper.MoveNumberedMultiple(sNumbered, iDir, plainsList.ListItems.Count - 1)
                If PlainsWin.isLoaded Then
                    Call PlainsWin.Toolwindow.ListHelper.MoveNumberedMultiple(sNumbered, iDir, plainsList.ListItems.Count - 1)
                End If
                If r > 0 Then
                    Call zPlainsChanged
                End If
            End If
        
        
        Case "item_add":
            plainsList_AddFileDialog
            
        Case "item_clear":
            If plainsList.ListItems.Count > 0 Then
                If MsgBox("Remove all wordlist file(s)?", vbQuestion + vbYesNo + vbDefaultButton1, "Clear wordlists") = vbYes Then
                    plainsList.ListItems.Clear
                    If PlainsWin.isLoaded Then
                        PlainsWin.Toolwindow.list.ListItems.Clear
                    End If
                    Call zPlainsChanged
                End If
            End If
        
        Case "item_del": 'delete selected item(s)
            If plainsListHelper.SelCount > 0 Then
                sNumbered = Me.plainsListHelper.NumberedSelected
                Call Me.plainsListHelper.RemoveNumbered(sNumbered)
                If PlainsWin.isLoaded Then
                    Call PlainsWin.Toolwindow.ListHelper.RemoveNumbered(sNumbered)
                End If
                Call zPlainsChanged 'signal changes
            End If

        Case "list_mode":
            plainsList.View = lvwReport - plainsToolbar.Buttons.Item("list_mode").Value
            
        Case "select_all":
            plainsListHelper.SelectAll
            
        Case "select_invert":
            plainsListHelper.SelectInverse
            
        Case Else:
            Stop
            zPlainsListCmd = 0
           
    End Select

End Function








'return a random caption for the go/start/hash button
Private Function zRandomCaption() As String
Dim Caption As String

    Randomize
    Select Case Int(Rnd * 5)
        Case 0:
            Caption = "Gimme cr&ackers"
        Case 1:
            Caption = "Power of the &Atom"
        Case 2:
            Caption = "I am the H&ashkiller!"
        Case 3:
            ' cmdGo(0).Caption = "Ich will einen H&ash erhaschen..."
            Caption = "I want to catch a h&ash ..."
        Case 4:
            'cmdGo(0).Caption = "H&ash mich, ich bin ein Digest"
            Caption = "H&ash me, I'm a digest."
    End Select

zRandomCaption = Caption
End Function


Private Sub bruteLenText_GotFocus(Index As Integer)
    Call textbox_select_all(bruteLenText(Index))
End Sub

Private Sub hashSeperator_GotFocus()
    Call textbox_select_all(hashSeperator)
End Sub

Private Sub recoveryModeCombo_Change()

    zOptFramesChange
    Job_Change
    
End Sub

Private Sub recoveryModeCombo_Click()
    recoveryModeCombo_Change
End Sub


Private Sub bruteCharsText_LostFocus()

    'recent handler
    Me.BruteChars.RecentTouch
    
End Sub

Private Sub bruteLenText_Change(Index As Integer)
    Job_Change
End Sub

Private Sub bruteLenText_LostFocus(Index As Integer)
Dim iVal As Long

    iVal = Int(Val(bruteLenText(Index).Text))
    If iVal < 1 Then iVal = 1
    
    If Index = 0 Then
        If iVal > Int(Val(bruteLenText(1).Text)) Then
            iVal = Int(Val(bruteLenText(1).Text))
        End If
    Else
        If iVal < Int(Val(bruteLenText(0).Text)) Then
            iVal = Int(Val(bruteLenText(0).Text))
        End If
    End If
    
    bruteLenText(Index).Text = CStr(iVal)
    
End Sub


Private Sub cmdGo_Click(Index As Integer)
    
    Select Case Index
        Case 0:
            Me.ExecuteHashcat
            If outFileCheck.Value = 1 And viewerCheck.Value = 1 Then
                Call textbox_fileedit(outFileText, """" & HCGUI_directory(0) & "outfileviewer.exe""")
            End If
            
    End Select

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub commandlineDisplayMenu_Click(Index As Integer)
    Me.commandlineDisplay = Index
End Sub

Private Sub commandlineMenu_Click(Index As Integer)
Dim cItem As MSComctlLib.Panel

    Select Case Index
        Case 0: 'copy to clipboard
            Clipboard.Clear
            Call Clipboard.SetText(Me.statusBar.Panels(1).Text)
        
        Case 1 'show toggle
            Set cItem = statusBar.Panels.Item(1)
            Call statusBar_PanelDblClick(cItem)
        
    End Select
End Sub


Private Sub bruteCharsText_Change()
    Job_Change
End Sub

Private Sub Form_Initialize()
    
    InitCommonControls 'XP Manifest / theme  support

    m_RuleMode = -1 'defaults to -1 to detect changes
    Set Me.RecentJobs = New cRecent
    
    Set plainsListHelper = New cLvHelper
    Set plainsListHelper.list = Me.plainsList
    
    Set Me.CommandlineWin = newTwHelper(Me, New fCommandline)
    
    Set Me.DrhashWin = New cTwHelper
    DrhashWin.Init Me, New fHash
    
    Set Me.PlainsWin = New cTwHelper
    PlainsWin.Init Me, New fPlains
    
    zFormMenuinit
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

     If Shift = 2 Then
        Select Case KeyCode
            Case 78: If menuFileOp(0).Enabled Then Call menuFileOp_Click(0)
            Case 79: If menuFileOp(1).Enabled Then Call menuFileOp_Click(1)
            Case 83: If menuFileOp(4).Enabled Then Call menuFileOp_Click(4)
            Case 84: If menuFileOp(5).Enabled Then Call menuFileOp_Click(5)
            Case 87: If menuFileOp(7).Enabled Then Call menuFileOp_Click(7)
        End Select
    End If
End Sub


Private Sub Form_Load()
Dim i As Long

    'have fun with the button
    cmdGo(0).Caption = zRandomCaption()
    
    ' bind recent file list
    Call m_RecentMenu.Init(Me.RecentJobs, Me.menuFileRecent, Me.menuFileOp.Item(2))
    
    ' bind cmd and recents to file textboxes / combos
    Set HashFile = New cFbHelper
    Call HashFile.Init(Me.hashFileText, Me.hashFileTb, True)
    
    Set BruteChars = New cFbHelper
    Call BruteChars.Init(Me.bruteCharsText, Nothing, True, False)
    
    Set OutFile = New cFbHelper
    Call OutFile.Init(Me.outFileText, Me.outFileTb, True)
    
    Set RuleFile = New cFbHelper
    Call RuleFile.Init(Me.ruleFileText, Me.ruleFileTB, True)
    
    Set SaltFile = New cFbHelper
    Call SaltFile.Init(Me.saltFileText, Me.saltFileTb, True)
    
    
    ' initialize statusbar (and recovery modes combo)
    Me.BinaryOs = HCGUI_BinOs
    Me.Binary = HCGUI_bin().Path 'hashcat
    statusBar.Panels.Item(3).Text = App.Major & "." & App.Minor & "." & App.Revision & " " 'gui
    statusBar.Style = sbrNormal
    
    ' fill hashmodes
    Call zFillHashmodes(Me.hashModeCombo)
    
    'load ini
    Call ini_load(HCGUI_Inifile)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        ' X pressed
        If Not Me.FileAskQuit Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub Form_Resize()

   On Error Resume Next
   headerPic.Top = 2 * Screen.TwipsPerPixelY
   fraSep(0).Move -6 * Screen.TwipsPerPixelX, headerPic.Top + headerPic.Height - 3 * Screen.TwipsPerPixelY, Me.ScaleWidth + 12 * Screen.TwipsPerPixelX
   fraSep(2).Move -6 * Screen.TwipsPerPixelX, -3 * Screen.TwipsPerPixelY, Me.ScaleWidth + 12 * Screen.TwipsPerPixelX
   fraSep(1).Move -6 * Screen.TwipsPerPixelX, fraSep(1).Top, Me.ScaleWidth + 12 * Screen.TwipsPerPixelX

End Sub

Private Sub Form_Terminate()
    Set Me.RecentJobs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Me.DrhashWin = Nothing
    
    Call ini_save(HCGUI_Inifile)
    
End Sub
Private Sub hashFile_Changed()
    Job_Change
End Sub
Private Sub hashFile_Click(sKey As String)

    Select Case sKey
        Case "cmd_browse":
            If textbox_filebrowse(hashFileText, , _
                        "All Files (*.*)|*.*|Hash Files (*hash*)|*hash*|PasswordPro files (*.Hashes)|(*.Hashes)|Text Files (*.txt)|*.txt", _
                        "Select Hashfile", "txt", OFN_HideReadOnly) Then
                HashFile.RecentTouch
            End If
        Case "cmd_edit":
            Call textbox_fileedit(hashFileText)
        Case "cmd_paste":
            Dim fPasteFrm As New fPaste
            Call fPaste.OpenForm(Me)
            Unload fPasteFrm
            Set fPasteFrm = Nothing
            
    End Select

End Sub




Private Sub hashFileTb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call HashFile.TriggerClick(Button.Key)
End Sub

Private Sub hashFileText_Change()
   
    Call HashFile.Trigger(ChangeEvent)
    
End Sub
Private Sub hashFileText_Click()

    Call HashFile.Trigger(ClickEvent)
    
End Sub
Private Sub hashFileText_LostFocus()
    HashFile.RecentTouch
End Sub
Private Sub hashFileText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If oledd_is_files(Data) Then
        hashFileText.Text = Data.Files.Item(1)
        textbox_select_all hashFileText
    End If
End Sub
Private Sub hashFileText_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4 'vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub



Private Sub hashModeCombo_Change()

    zOptFramesChange
    Job_Change
    
End Sub

Private Sub hashModeCombo_Click()

    hashModeCombo_Change
    
End Sub

Private Sub hashModeToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim p1 As POINTAPI

    If Me.DrhashWin.isLoaded Then
        Me.DrhashWin.Toggle
    Else
        'below the button:
        'p1 = HCGUI_Control_RightLeft(Me.hashModeCommand)
        'FIXME remove obsolete Call SetWindowLong(DrhashWin.Toolwindow.hwnd, GWL_HWNDPARENT, Me.hwnd)
        'DrhashWin.Toolwindow.Move p1.x * Screen.TwipsPerPixelX - DrhashWin.Toolwindow.Width, p1.y * Screen.TwipsPerPixelY
        
        'centered on window
        DrhashWin.Toolwindow.Move (Me.Width - DrhashWin.Toolwindow.Width) / 2 + Me.Left, (Me.Height - DrhashWin.Toolwindow.Height) / 2 + Me.Top
        
        Me.DrhashWin.Toggle
    End If

End Sub

Private Sub hashSeperator_Change()
    Job_Change
End Sub

Private Sub helpMenuEntry_Click(Index As Integer)
' FIXME Dim oExec As New cExec
' FIXME Dim sCmd As String
' FIXME Dim sTitle As String
Dim r As Long
Dim cShellExec As New cShellExec

    Select Case Index
        Case 0: 'info
            Call fInfo.ShowInfo(Me, 0)
            Exit Sub
            
        Case 1: '---
            
        Case 2: 'help
            Call fInfo.ShowInfo(Me, 1)
            Exit Sub
            
        Case 3: 'EULA
            Call fInfo.ShowInfo(Me, 2)
            Exit Sub
            
        Case 4: 'Online Forum
            r = cShellExec.Exec(Me.hWnd, "open", "http://hashcat.net/forum/", "", "", ShellExec_SW_SHOWNORMAL)
            Exit Sub
            
        Case 6: 'DebugInfo
            Call fInfo.ShowInfo(Me, 3, zDebugInfo(), "Debug Info")
            Exit Sub
            
        Case Else:
            MsgBox "Your command was undefined.", vbCritical, "Critical Programflow"
            Exit Sub
    End Select
    

End Sub






Private Sub menuFile_Click(Index As Integer)

    Select Case Index
        Case 0: 'consle stay open ftw
            Me.menuFile.Item(Index).Checked = Not Me.menuFile.Item(Index).Checked
            
        Case 1: 'view resource settings ftw
            Me.menuFile.Item(Index).Checked = Not Me.menuFile.Item(Index).Checked
            resourcesFrame.Visible = Abs(Me.menuFile.Item(Index).Checked)
            
        Case 2: '---
        
        Case 3: 'select hashcat binary
            Call fSettings.Ask(Me)
            
    End Select
    
End Sub


Private Sub menuFileDebug_Click(Index As Integer)
Dim r As Long
Dim runForm As fRun

    Select Case Index
        Case 0: '---
        
        Case 1: 'Run
            Set runForm = New fRun
            fRun.Show 0, Me
            Set fRun = Nothing
        
        Case 2: 'shell
            On Error Resume Next
                r = Shell("cmd", vbNormalFocus)
                Me.Caption = r
            On Error GoTo 0
            
        Case 3: 'winconsole
            On Error Resume Next
                r = Shell("wineconsole cmd", vbNormalFocus)
                Me.Caption = r
            On Error GoTo 0
            
    End Select
    

End Sub

Private Sub menuFileOp_Click(Index As Integer)
Dim cc As cCommonDialog
Dim sFile As String
Dim iRet As Long

    Select Case Index
        Case 0: 'new
            If Not Me.FileAskDirty("New hashcat Job") Then Exit Sub
            Set Me.Job = New cJob
            m_Filename = ""
            Me.Dirty = False
            
        Case 1: 'open
            If Not Me.FileAskDirty("Open new hashcat Job") Then Exit Sub
            
            'offer dialog
            Set cc = New cCommonDialog
            If cc.VBGetOpenFileName(sFile, , , , , , "hashcat Job (.hcj)|*.hcj|All Files|*.*", , , "Open hashcat Job", "hcj", Me.hWnd, OFN_HideReadOnly) Then
                'load file
                Call Me.FileOpen(sFile)
            End If
            
        Case 4: 'save
            If m_Filename = "" Then
                'save as action
                Call menuFileOp_Click(5)
            Else
                'save file
                Call HCGUI_job_to_ini(Me.Job, m_Filename)
                'set dirty flag
                Me.Dirty = False
            End If
            
            
        Case 5: 'save as
            'offer dialog
            Set cc = New cCommonDialog
            If cc.VBGetSaveFileName(sFile, , , "hashcat Job (.hcj)|*.hcj|All Files|*.*", , HCGUI_directory(100), "Save hashcat Job", "hcj", Me.hWnd, OFN_HideReadOnly) Then
                'save file
                m_Filename = sFile
                Call HCGUI_job_to_ini(Me.Job, m_Filename)
                Me.RecentJobs.Touch m_Filename
                'set dirty flag
                Me.Dirty = False
            End If
            
        Case 7: 'exit
            If Me.FileAskQuit Then
                Unload Me
            End If
                        
    End Select
    
End Sub
Private Sub menuFileRecent_Click(Index As Integer)
Dim sFile As String

    sFile = menuFileRecent.Item(Index).Tag
    
    ' check for dirtyness
    If Not Me.FileAskDirty("Open new hashcat Job") Then
        Exit Sub
    End If
    
    ' open the new file
    Call Me.FileOpen(sFile)

End Sub

Private Sub menuView_Click(Index As Integer)

    Select Case Index
        Case 0: 'wordlist manager
            zPlainsEditorCmd
            menuView(Index).Checked = PlainsWin.Visible
        
        Case 1: 'dr hash
            Call hashModeToolbar_ButtonClick(Me.hashModeToolbar.Buttons.Item("drhash"))
            menuView(Index).Checked = DrhashWin.Visible
        
        Case 2: 'commandline
            Call commandlineMenu_Click(1)
            menuView(Index).Checked = CommandlineWin.Visible
        
    End Select
    
End Sub

Private Sub menuViewParent_Click()
    
    menuView(0).Checked = PlainsWin.Visible
    menuView(1).Checked = DrhashWin.Visible
    menuView(2).Checked = CommandlineWin.Visible

End Sub

Private Sub OutFile_Changed()
    Job_Change
End Sub

Private Sub OutFile_Click(sKey As String)

    Select Case sKey
        Case "cmd_browse":
            If textbox_filebrowse(outFileText, False, _
                        "All Files (*.*)|*.*|Outfiles (*.out)|*.out|Text Files (*.txt)|*.txt", _
                        "Select Outfile", "out", OFN_HideReadOnly Or OFN_PathMustExist Or OFN_CREATEPROMPT) Then
                OutFile.RecentTouch
            End If

        Case "cmd_edit":
            Call textbox_fileedit(outFileText, """" & HCGUI_directory(0) & "outfileviewer""")
        
    End Select

End Sub


Private Sub outFileCheck_Click()
    Job_Change
End Sub

Private Sub outFileTb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call OutFile.TriggerClick(Button.Key)
End Sub

Private Sub outFileText_Change()

    Call OutFile.Trigger(ChangeEvent)
    
End Sub

Private Sub outFileText_LostFocus()

    OutFile.RecentTouch
    
End Sub

Private Sub outFileText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If oledd_is_files(Data) Then
        outFileText.Text = Data.Files.Item(1)
        textbox_select_all outFileText
    End If
End Sub
Private Sub outFileText_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4 'vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub






Private Sub plainsList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim keystat(0 To 255) As Byte ' receives key status information for all keys
Dim retval As Long ' return value of function
Dim sNumbered As String

    retval = GetKeyboardState(keystat(0)) ' In VB, the array is passed by referencing element #0.
    
    'if ctrl is pressed, check/de-check all selected
    If (keystat(&H11) And &H80) = &H80 Then
        sNumbered = plainsListHelper.NumberedSelected
        Call plainsListHelper.CheckNumbered(sNumbered, Item.Checked)
        If PlainsWin.isLoaded Then
            Call PlainsWin.Toolwindow.ListHelper.CheckNumbered(sNumbered, Item.Checked)
        End If
    Else
        If PlainsWin.isLoaded Then
            PlainsWin.Toolwindow.list.ListItems.Item(Item.Index).Checked = Item.Checked
        End If
    End If
    Call zPlainsChanged

End Sub

Private Sub plainsList_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not plainsList.SelectedItem Is Nothing Then
        'FIXME tooltiptext is normally not need to change if listview is in detailed mode -> the code below creates sideeffects then
        If plainsList.View = lvwReport Then
            plainsList.ToolTipText = ""
        Else
            'plainsList.ToolTipText = plainsList.SelectedItem.ToolTipText
        End If
    End If
    
End Sub

Private Sub plainsList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r As Long

    If Shift = 0 Then
        Select Case KeyCode
            Case 45: r = zPlainsListCmd("item_add")
            Case 46: r = zPlainsListCmd("item_del")
            Case 93: PopupMenu Me.plainsMenu
            Case 123: r = zPlainsListCmd("item_clear")
            Case Else:
                'Stop
        End Select
    ElseIf Shift = 2 Then
        Select Case KeyCode
            Case 65: r = zPlainsListCmd("select_all")
            Case 67: r = zPlainsListCmd("cb_copy")
            Case 73: r = zPlainsListCmd("select_invert")
            Case 86: r = zPlainsListCmd("cb_paste")
            Case 88: r = zPlainsListCmd("cb_cut")
        End Select
    ElseIf Shift = 3 Then
        Select Case KeyCode
            Case 73: r = zPlainsListCmd("check_invert")
            Case 77: r = zPlainsListCmd("check_all")
            Case 85: r = zPlainsListCmd("check_none")
        End Select
    ElseIf Shift = 4 Then
        Select Case KeyCode
            Case 35: r = zPlainsListCmd("item_tobottom")
            Case 36: r = zPlainsListCmd("item_totop")
            Case 40: r = zPlainsListCmd("item_down")
            Case 38: r = zPlainsListCmd("item_up")
        End Select
    End If

End Sub

Private Sub plainsList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu Me.plainsMenu
    End If
End Sub

Private Sub plainsList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
    If oledd_is_files(Data) Then
        For i = 1 To Data.Files.Count
            Call plainsList_AddFile(Data.Files.Item(i))
        Next
    End If
End Sub

Private Sub plainsList_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4
    Else
        Effect = vbDropEffectNone
    End If
End Sub


Private Sub plainsMenuEntry_Click(Index As Integer)
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
        r = zPlainsListCmd(sCmd)
    End If

End Sub


' set atextboxs content based on a file selection
Private Function textbox_filebrowse(ctlTextbox As Control, _
                                    Optional FileMustExist As Boolean = True, _
                                    Optional Filter As String = "All (*.*)| *.*", _
                                    Optional DlgTitle As String, _
                                    Optional DefaultExt As String, _
                                    Optional flags As Long = 0) As Boolean
Dim oFi As New cFileinfo
Dim oCc As New cCommonDialog
Dim sFile As String
Dim Owner As Long
Dim InitDir As String

    oFi.Path = ctlTextbox.Text
    Owner = ctlTextbox.Parent.hWnd
    
    InitDir = oFi.ExistingDir
    
    'default directory
    If InitDir = "" Then
        Select Case ctlTextbox.Name
        
            Case "hashFileText":
                InitDir = HCGUI_directory(100)

            Case "ruleFileText":
                InitDir = HCGUI_directory(102)
                
            Case "saltFileText":
                InitDir = HCGUI_directory(103)

            Case Else:
                InitDir = HCGUI_directory(100)
        End Select
    End If
    
    If oCc.VBGetOpenFileName(sFile, , , , , , Filter, , InitDir, DlgTitle, DefaultExt, Owner, flags) Then
        ctlTextbox.Text = sFile
        Call textbox_select_all(ctlTextbox)
        ctlTextbox.SetFocus
        textbox_filebrowse = True
    End If

End Function

'edit file based on textbox text
Private Function textbox_fileedit(ctlTextbox As Object, Optional sCmd As String = "cmd /c start wordpad")
Dim sFile As String
Dim iMousepointer As Long
Dim r As Long

    sFile = ctlTextbox.Text
    
    iMousepointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    ' r = Shell("cmd /c """ & sCmd & " " & paramsafe(sFile, Windows) & " & pause""", vbHide)
    On Error Resume Next
    r = Shell(sCmd & " " & paramsafe(sFile, Windows), vbNormalFocus)
    If Err <> 0 Then
        MsgBox "Error launching external application."
    End If
    On Error GoTo 0
    
    Screen.MousePointer = iMousepointer
    
End Function
'
' helper to be called if Plains changed (on this form)
'
Private Sub zPlainsChanged()
    'change the job
    Job_Change
End Sub
Private Sub RecentFiles_Update(Index As Long)
End Sub



Private Sub RecentJobs_Update()
    m_RecentMenu.refresh
End Sub

Private Sub RuleFile_Changed()
    Job_Change
End Sub

Private Sub RuleFile_Click(sKey As String)

    Select Case sKey
        Case "cmd_browse": 'browse
            If textbox_filebrowse(ruleFileText, , _
                                    "Ruleset files (*.rule)|*.rule|All Files (*.*)|*.*", _
                                    "Select Ruleset file", "rule", OFN_HideReadOnly) Then
                RuleFile.RecentTouch
            End If
            
        Case "cmd_edit": 'edit
            Call textbox_fileedit(ruleFileText)
        
    End Select
    
End Sub


Private Sub ruleFileTB_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call RuleFile.TriggerClick(Button.Key)
End Sub

Private Sub ruleFileText_Change()

    'activate the according option
    Me.ruleOption(1).Value = True
    
    'trigger
    Call RuleFile.Trigger(ChangeEvent)
    
End Sub

Private Sub ruleFileText_Click()

    Call RuleFile.Trigger(ClickEvent)
    
End Sub


Private Sub ruleFileText_LostFocus()
    
    'recent handler
    RuleFile.RecentTouch
    
End Sub

Private Sub ruleFileText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If oledd_is_files(Data) Then
        ruleFileText.Text = Data.Files.Item(1)
        textbox_select_all ruleFileText
    End If
End Sub
Private Sub ruleFileText_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4 'vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub


Private Sub ruleGenerateText_Change()
Dim s As String

    ' nicely formatted value for the caption (usability)
    s = Format(Val(Me.RuleGenerate), "Standard")
    zLbl(0).Caption = Left(s, Len(s) - 3)
    
    ' activate the according option
    Me.ruleOption(2).Value = True
    
    Job_Change
    
End Sub


Private Sub ruleGenerateText_Click()

    ruleGenerateText_Change
    
End Sub

Private Sub ruleGenerateText_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case &H30 To &H39:
        Case &H20:
        Case Asc("k"):
        Case Asc("K"):
        Case Asc("m"):
        Case Asc("M"):
        Case 8:
        Case Else
            KeyAscii = 0
    End Select

End Sub


Private Sub ruleOption_Click(Index As Integer)
    
    If m_RuleMode <> Index Then
        Job_Change
    End If
    
    m_RuleMode = Index

End Sub

Private Sub SaltFile_Changed()
    Job_Change
End Sub

Private Sub SaltFile_Click(sKey As String)

    Select Case sKey
        Case "cmd_browse":
            If textbox_filebrowse(saltFileText, False, _
                        "Saltiles (*.salt)|*.salt|Text Files (*.txt)|*.txt|All Files (*.*)|*.*", _
                        "Select Saltfile", "salt", OFN_HideReadOnly Or OFN_PathMustExist Or OFN_CREATEPROMPT) Then
                SaltFile.RecentTouch
            End If
            
        Case "cmd_edit":
            Call textbox_fileedit(saltFileText)
            
    End Select
    
End Sub

Private Sub saltFileCheck_Click()
    Job_Change
End Sub

Private Sub saltFileTb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call SaltFile.TriggerClick(Button.Key)
End Sub

Private Sub saltFileText_Change()

     Call SaltFile.Trigger(ChangeEvent)
     
End Sub

Private Sub saltFileText_LostFocus()

    SaltFile.RecentTouch
    
End Sub


Private Sub saltFileText_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If oledd_is_files(Data) Then
        saltFileText.Text = Data.Files.Item(1)
        textbox_select_all saltFileText
    End If
End Sub


Private Sub saltFileText_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If oledd_is_files(Data) Then
        Effect = 4 'vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub segmentText_Change()
    Job_Change
End Sub

Private Sub segmentText_GotFocus()
    Call textbox_select_all(Me.ActiveControl)
End Sub


Private Sub skipText_Change(Index As Integer)
    Job_Change
End Sub
Private Sub skipText_GotFocus(Index As Integer)
    Call textbox_select_all(Me.ActiveControl)
End Sub


Private Sub skipText_LostFocus(Index As Integer)
Dim vValue As Variant

    vValue = Cnv_Str2Dec(skipText(Index).Text)
    If TypeName(vValue) = "String" Then
        skipText(Index).Text = ""
    ElseIf TypeName(vValue) = "Decimal" Then
        If vValue < 0 Then
            skipText(Index).Text = "0"
        ElseIf vValue > CDec("18446744073709551615") Then
            vValue = CDec("18446744073709551615")
            skipText(Index).Text = CStr(vValue)
        End If
    End If

End Sub

Private Sub statusBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu Me.commandlineMenuParent
    End If
    
End Sub

Private Sub statusBar_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
Dim p1 As POINTAPI

    Select Case Panel.Index
        Case 1:
            If Not Me.CommandlineWin.isLoaded Then
                Me.CommandlineWin.Load
                Commandline_Change 'update window
                'default pos
                p1 = POS_control_topleft(Me.statusBar)
                CommandlineWin.Toolwindow.Move p1.x * Screen.TwipsPerPixelX, (p1.y * Screen.TwipsPerPixelY) - CommandlineWin.Toolwindow.Height
                              
            End If
            Me.CommandlineWin.Toggle
            commandlineMenu(1).Checked = Me.CommandlineWin.Visible
            
        Case 2:
            Call menuFile_Click(3)
        
        Case Else:
            Call helpMenuEntry_Click(0)
    End Select

End Sub


Private Sub threadsText_Change()
    Job_Change
End Sub

Private Sub plainsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim r As Long

    r = zPlainsListCmd(Button.Key)
    
End Sub


Private Sub threadsText_GotFocus()
        Call textbox_select_all(Me.ActiveControl)
End Sub

Private Sub toggleLenText_Change(Index As Integer)
    Job_Change
End Sub

Private Sub toggleLenText_GotFocus(Index As Integer)
    Call textbox_select_all(toggleLenText(Index))
End Sub


Private Sub toggleLenText_LostFocus(Index As Integer)
Dim iVal As Long

    iVal = Int(Val(toggleLenText(Index).Text))
    If iVal < 1 Then iVal = 1
    
    If Index = 0 Then
        If iVal > Int(Val(toggleLenText(1).Text)) Then
            iVal = Int(Val(toggleLenText(1).Text))
        End If
    Else
        If iVal < Int(Val(toggleLenText(0).Text)) Then
            iVal = Int(Val(toggleLenText(0).Text))
        End If
    End If
    
    toggleLenText(Index).Text = CStr(iVal)
End Sub

Private Sub zLbl_Click(Index As Integer)

    Select Case Index
        Case 1:
            bruteCharsText.SetFocus
        Case 2, 3:
            bruteLenText(Index - 2).SetFocus
        Case 4, 5:
            toggleLenText(Index - 4).SetFocus
        Case 6:
            threadsText.SetFocus
        Case 7, 8:
            segmentText.SetFocus
        Case 9, 10:
            skipText(Index - 9).SetFocus
        Case 11:
            recoveryModeCombo.SetFocus
        Case 12:
            hashModeCombo.SetFocus
        Case 13:
            hashFileText.SetFocus
        Case 14:
            hashSeperator.SetFocus
        Case 15:
              plainsList.SetFocus
        
    End Select
    
End Sub


