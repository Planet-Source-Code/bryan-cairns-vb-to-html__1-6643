VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Source Documentor"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5535
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9763
      _Version        =   393217
      BackColor       =   16777215
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog commdlg 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0978
            Key             =   "unread"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AE0
            Key             =   "writing"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C3C
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D98
            Key             =   "error"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6376
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4080
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
            Picture         =   "Form1.frx":0EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Document"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   4575
         TabIndex        =   3
         Top             =   0
         Width           =   4575
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   25
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Text            =   "Status:"
            TextSave        =   "Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9790
            Text            =   "Idle"
            TextSave        =   "Idle"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuzz1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDoc 
      Caption         =   "Document"
      Begin VB.Menu mnudo 
         Caption         =   "Do it now"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
ListView1.ListItems.Clear
sLastFile = ""
ProjectName = ""
ProjectVersion = ""
Prodir = ""
ProType = ""
ProMinVer = ".0"
'RichTextBox1.SelColor = vbBlack
'RichTextBox1.SelColor = &H800000
'RichTextBox1.SelColor = &H8000&

End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.Width - 120
ListView1.Height = Me.Height - 1400
Picture1.Width = Me.Width - Picture1.Left - 120
ProgressBar1.Width = Picture1.Width - 75
ListView1.ColumnHeaders(1).Width = ListView1.Width - 350
End Sub

Private Sub mnudo_Click()
Dim sPath As String
If ListView1.ListItems.Count > 0 Then
    sPath = BrowseForFolder(Me.hWnd, "Select a Folder to Write to", "c:\")
    If Len(sPath) > 0 Then
    'do the conversion
    DocDir = sPath & "\"
        Select Case ProType
        Case Is = ".vbp" 'visual basic
        Call InitColorize ' for the html to rtf coloring
        WriteVBHTMLFiles
        Case Is = ".dpr" 'delphi
        'comming soon?
        Case Is = "dsw" 'visual c++
        'comming soon?
End Select
    End If

End If

End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNew_Click()
ListView1.ListItems.Clear
sLastFile = ""
ProjectName = ""
ProjectVersion = ""
Prodir = ""
ProType = ""
ProMinVer = ".0"
End Sub

Private Sub mnuOpen_Click()
'show the comdlg
On Error GoTo cCan
Dim sEXT As String
commdlg.FileName = ""
commdlg.Filter = "Visual Basic Project (*.vbp)|*.vbp|Delphi Project (*.dpr)|*.dpr|Visual C++ Project (*.dsw)|*.dsw"
commdlg.FilterIndex = 1
commdlg.Flags = cdlOFNFileMustExist + cdlOFNExplorer
commdlg.DefaultExt = ".vbp"
commdlg.CancelError = True
'commdlg.InitDir = App.Path
commdlg.ShowOpen

'select case here
sEXT = ParsePath(commdlg.FileName, 3)
MousePointer = 11
ProType = LCase(sEXT)
Select Case ProType
Case Is = ".vbp"
ParseProjectFileVB commdlg.FileName
MousePointer = 0
Case Is = ".dpr"
Case Is = "dsw"
End Select
'DocDir = ParsePath(commdlg.FileName, 0) & ParsePath(commdlg.FileName, 1) & "Documentation\"
Exit Sub
cCan:
If Err <> cdlCancel Then
MsgBox Err.Description
MousePointer = 0
End If
Exit Sub
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
Case Is = 1 'new
mnuNew_Click
Case Is = 2 'open
mnuOpen_Click
Case Is = 3 'document
mnudo_Click
End Select
End Sub
