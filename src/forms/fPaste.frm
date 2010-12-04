VERSION 5.00
Begin VB.Form fPaste 
   Caption         =   "UseHashes from Clipboard"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "Bitstream Vera Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPaste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7395
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox hashesText 
      BeginProperty Font 
         Name            =   "Bitstream Vera Sans Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Frame cmdFrame 
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   4380
      TabIndex        =   0
      Top             =   4440
      Width           =   2355
      Begin VB.CommandButton cmdButton 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   1095
      End
   End
End
Attribute VB_Name = "fPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
   (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
      
Private Declare Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
    
Private Const MAX_PATH As Long = 260

Private m_Parent As Form

Public Sub OpenForm(fParent As Form)
Dim sText As String

    Set m_Parent = fParent
    
    sText = zNormalizeText(Clipboard.GetText(vbCFText), vbCrLf)
    
    Me.hashesText.SelText = sText
    
    Me.Icon = fParent.Icon
    Me.Show 1, fParent
    
End Sub

Public Property Get Parent() As fMain
    Set Parent = m_Parent
End Property


Private Sub cmdButton_Click(Index As Integer)

    Select Case Index
        Case 0:
            Call zDoTextImport
    End Select
    
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Hide
    
End Sub

Private Sub Form_Resize()
Dim h As Long

    h = Me.ScaleHeight - cmdFrame.Height
    If h < 0 Then h = 0
    
    hashesText.Move 0, 0, Me.ScaleWidth, h
    cmdFrame.Move Me.ScaleWidth - cmdFrame.Width, Me.ScaleHeight - cmdFrame.Height
    
End Sub


Private Sub zDoTextImport()
Dim r As Long
Dim sBuffer As String, iBufferLen As Long
Dim sTempDir As String
Dim sFile As String
Dim sHashes As String
Dim iFn As Long

    iBufferLen = 1024
    sBuffer = Space(iBufferLen)
    r = GetTempPath(iBufferLen, sBuffer)
    If r = 0 Then
        MsgBox "Failed. Error while determing temporary directory."
        Exit Sub
    End If
    
    If r > iBufferLen Then
        MsgBox "Failed. Temporary directory path is too long."
        Exit Sub
    End If
    
    sTempDir = Left(sBuffer, r)
    
    sBuffer = Space(iBufferLen)
    
    sFile = Space$(MAX_PATH)
    r = GetTempFileName(sTempDir, "hash", 0, sFile)
    
    If r = 0 Then
        MsgBox "Faled. Error while creating temporary filename"
        Exit Sub
    End If
    
    sFile = Left$(sFile, InStr(sFile, Chr$(0)) - 1)
            
    sHashes = zNormalizeText(Me.hashesText.Text)
    
    iFn = FreeFile()
    Open sFile For Output As iFn
        Print #iFn, sHashes;
    Close iFn
    
    Me.Parent.hashFileText.Text = sFile
End Sub
Private Function zNormalizeText(ByVal sText As String, Optional sLineEnding As String = vbLf) As String
Dim aText() As String
Dim iCount As Long, i As Long
Dim vText As Variant

    If Len(sText) Then
        aText = Split(sText, vbCrLf)
        iCount = UBound(aText)
        For i = 0 To iCount
            aText(i) = Trim(aText(i))
        Next i
        sText = Join(aText, sLineEnding)
    End If
    
    zNormalizeText = sText
End Function

