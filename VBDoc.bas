Attribute VB_Name = "VBDoc"

'This Module is for converting VB to HTML with color coding





Public Sub ParseProjectFileVB(sfile As String)
'open the file and parse the lines
On Error GoTo EH
Dim TextLine
Dim sSTART As String
Dim thefile As String
Dim ITMX As ListItem
Dim I As Integer
Dim sFileDir As String
Dim sFileDrive As String
Dim ipos As Integer
sLastFile = sfile
I = 1
Form1.ProgressBar1.Min = I
Form1.ProgressBar1.Max = FileLines(sfile)
Form1.ProgressBar1.Value = Form1.ProgressBar1.Min
sFileDrive = ParsePath(sfile, 0)
sFileDir = ParsePath(sfile, 1)
sFileDrive = sFileDrive & sFileDir
Form1.ListView1.ListItems.Clear
Open sfile For Input As #1
Do While Not EOF(1)
Form1.ProgressBar1.Value = I
Form1.StatusBar1.Panels(2).Text = "Reading Line " & I
  I = I + 1
   Line Input #1, TextLine
'get the start of the line
ipos = InStr(1, TextLine, "=", vbBinaryCompare)
If ipos <> 0 Then
sSTART = Mid(TextLine, 1, ipos - Len("="))
'Set ITMX = Form1.ListView1.ListItems.Add(, , sSTART, , 1)
thefile = Mid(TextLine, ipos + Len("="), Len(TextLine) - ipos)
    Select Case LCase(sSTART)
    Case Is = "form", "usercontrol", "module"
        ipos = InStr(1, TextLine, ";", vbBinaryCompare)
        If ipos <> 0 Then
        thefile = Mid(TextLine, ipos + Len(";") + 1, Len(TextLine) - ipos)
        Set ITMX = Form1.ListView1.ListItems.Add(, , sFileDrive & thefile, , 1)
        Else
        If ipos = 0 Then
        Set ITMX = Form1.ListView1.ListItems.Add(, , sFileDrive & thefile, , 1)
        End If
        End If
    Case Is = "name"
    ProjectName = Mid(thefile, 2, Len(thefile) - 2)
    Form1.Caption = Form1.Caption & " - " & ProjectName
    Case Is = "majorver"
    ProjectVersion = thefile
    Case Is = "minorver"
    ProMinVer = "." & thefile
    End Select
    End If

Loop
Close #1
Form1.ProgressBar1.Value = Form1.ProgressBar1.Min
Form1.StatusBar1.Panels(2).Text = Form1.ListView1.ListItems.Count & " Files in Project"
MousePointer = 0
Prodir = sFileDrive
Exit Sub
EH:
Close
Form1.ProgressBar1.Value = Form1.ProgressBar1.Min
Form1.ListView1.ListItems.Clear
MsgBox Err.Description
Exit Sub
End Sub

Public Sub WriteVBHTMLFiles()
'On Error Resume Next
'write the startup file
Dim sfile As String
Dim sWorkFile As String
Dim sDir As String
Dim sEXT As String
Dim I As Integer
Dim IFile As Integer
Dim Qoute As String
Qoute = Chr(34)
IFile = FreeFile
sDir = ParsePath(sLastFile, 0)
sfile = ParsePath(sLastFile, 1)
sWorkFile = DocDir 'sDir & sfile & "Documentation\"

sWorkFile = sWorkFile & "index.html"
Form1.ProgressBar1.Min = 1
Form1.ProgressBar1.Max = Form1.ListView1.ListItems.Count
Form1.ProgressBar1.Value = Form1.ProgressBar1.Min

    Open sWorkFile For Output As IFile
    Print #IFile, "<HTML><HEAD><TITLE>" & ProjectName & "</TITLE></HEAD>"
    Print #IFile, "<BODY BGCOLOR=""#FFFFFF"" Text=""#000000"" LINK=""#0000FF"" VLINK=""#000099"" ALINK=""#00FF00"">"
    Print #IFile, "</FONT><B><FONT SIZE=5 COLOR=""#008080""><P> " & ProjectName & " " & ProjectVersion & ProMinVer & " Source Code Documentation</P></B></FONT><FONT SIZE=2>"
    Print #IFile, "<P>Source Code Procedures and Function Broken down to a Form and Module level.<BR><BR><HR><BR>"

    For I = 1 To Form1.ListView1.ListItems.Count
    'write the start file
    'and all the other files
    Form1.ListView1.ListItems(I).SmallIcon = 3
    Form1.ProgressBar1.Value = I
        sEXT = ParsePath(Form1.ListView1.ListItems(I).Text, 3)
        sfile = ParsePath(Form1.ListView1.ListItems(I).Text, 2)
        'sfile = sfile & sEXT
    Print #IFile, "<a href=" & Qoute & sfile & ".html" & Qoute & ">" & sfile & sEXT & "</A><BR>" _
    & "<a href=" & Qoute & sfile & "src.html" & Qoute & ">Source</A><BR><BR>"
    'Form1.Caption = Prodir & "Documentation\" & sfile & ".txt"
    'write the source code to the file
    Form1.StatusBar1.Panels(2).Text = "Writing File " & I & " of " & Form1.ListView1.ListItems.Count
    Form1.ListView1.ListItems(I).SmallIcon = 3
    Form1.StatusBar1.Panels(2).Text = "Coloring File " & I & " of " & Form1.ListView1.ListItems.Count
    WriteTXTSCR Form1.ListView1.ListItems(I).Text, DocDir & sfile & "src.html", I
    Form1.StatusBar1.Panels(2).Text = "Documenting File " & I & " of " & Form1.ListView1.ListItems.Count
    WriteFRMSCR Form1.ListView1.ListItems(I).Text, DocDir & sfile & ".html", I
    Form1.StatusBar1.Panels(2).Text = "Finishing File " & I & " of " & Form1.ListView1.ListItems.Count
DoEvents
Next I
Print #IFile, "<BR></FONT></BODY></HTML>"
Close #IFile
Form1.ProgressBar1.Value = Form1.ProgressBar1.Min
 Form1.StatusBar1.Panels(2).Text = "Done"
End Sub



Private Sub WriteTXTSCR(sfile As String, WrFile As String, I As Integer)
'look for Attribute VB_Exposed = False
'Write the source code in color html
On Error GoTo EH
Form1.RichTextBox1.SelColor = vbBlack
Form1.RichTextBox1.Text = ""
Dim IFileA As Integer
Dim IFileB As Integer
Dim tmpSTR As String
Dim bFound As Boolean
Dim sEXT As String
Dim TextLine
tmpSTR = LCase("Attribute VB_Exposed = False")
bFound = False
IFileA = 2
IFileB = 3
If LCase(ParsePath(sfile, 3)) = ".bas" Then bFound = True

If CheckFile(sfile) = False Then
Form1.ListView1.ListItems(I).SmallIcon = 4
Exit Sub
End If
Open sfile For Input As IFileA 'source
'Open WrFile For Output As IFileB 'target
    If bFound = True Then
    Line Input #IFileA, TextLine
    End If
Do While Not EOF(IFileA)
   Line Input #IFileA, TextLine
   
If LCase(TextLine) = tmpSTR Then
bFound = True
Line Input #IFileA, TextLine
End If
If bFound = True Then
    'Print #IFileB, TextLine
    Form1.RichTextBox1.Text = Form1.RichTextBox1.Text & TextLine & vbCrLf
End If
Loop

'Close #IFileB
Close #IFileA
'we have all the text colorize and convert
Call ColorizeWords(Form1.RichTextBox1)

Form1.RichTextBox1.SaveFile WrFile, 0 'rtf
Form1.RichTextBox1.LoadFile WrFile, 1 'plain text

Open WrFile For Output As IFileB 'target
Print #IFileB, "<HTML><TITLE>Source Code</TITLE><BODY>" & RTF2HTML(Form1.RichTextBox1.Text) & "</BODY></HTML>"
Close #IFileB
Form1.RichTextBox1.Text = ""
Exit Sub
EH:
Form1.ListView1.ListItems(I).SmallIcon = 4
Exit Sub
End Sub

Private Sub WriteFRMSCR(sfile As String, WrFile As String, I As Integer)
'look for Attribute VB_Exposed = False
On Error GoTo EH
Dim IFileA As Integer
Dim IFileB As Integer
Dim tmpSTR As String
Dim bFound As Boolean
Dim sEXT As String
Dim TextLine
Dim bLook As Boolean
Dim ssName As String

ssName = ParsePath(sfile, 2)
ssName = ssName & ParsePath(sfile, 3)
bLook = False
bFound = False
IFileA = 2
IFileB = 3

If CheckFile(sfile) = False Then
Form1.ListView1.ListItems(I).SmallIcon = 4
Exit Sub
End If
Open sfile For Input As IFileA
Open WrFile For Output As IFileB
'write the htmlheader
Print #IFileB, "<HTML><HEAD><TITLE>" & ProjectName & "</TITLE></HEAD>"
Print #IFileB, "<BODY BGCOLOR=""#FFFFFF"" Text=""#000000"" LINK=""#0000FF"" VLINK=""#000099"" ALINK=""#00FF00"">"
Print #IFileB, "</FONT><B><FONT SIZE=5 COLOR=""#008080"">" & ssName & "</B></FONT><FONT SIZE=2>"
Print #IFileB, "<BR><BR><HR><BR>"
    
    
Do While Not EOF(IFileA)
    Line Input #IFileA, TextLine
    
    If ISWRITABLE(CStr(TextLine)) = True Then 'print the procedure
    Print #IFileB, "<br><FONT SIZE=3 COLOR=""#008080"">" & TextLine & "</B></FONT><br>"
    Do
        'look for the comments
        'advance a line
        Line Input #IFileA, TextLine
        tmpSTR = Left$(TextLine, 1)
        If tmpSTR = "'" Then 'it is a comment
        tmpSTR = Mid(TextLine, 2, Len(TextLine))
        Print #IFileB, tmpSTR & "<br>"
        Else
        If tmpSTR <> "'" Then
        Exit Do
        End If
        End If
    Loop
    
    End If
    
    
    


Loop

'write the htmlfooter
Print #IFileB, "<BR></FONT></BODY></HTML>"
Close #IFileB
Close #IFileA
Exit Sub
EH:
Form1.ListView1.ListItems(I).SmallIcon = 4
Exit Sub
End Sub

Private Function ISWRITABLE(sTMP As String) As Boolean
Dim ipos As Integer
Dim bFound As Boolean
bFound = False
ipos = InStr(1, LCase(sTMP), "private sub", vbBinaryCompare)
If ipos <> 0 Then bFound = True
ipos = InStr(1, LCase(sTMP), "public sub", vbBinaryCompare)
If ipos <> 0 Then bFound = True
ipos = InStr(1, LCase(sTMP), "private function", vbBinaryCompare)
If ipos <> 0 Then bFound = True
ipos = InStr(1, LCase(sTMP), "public function", vbBinaryCompare)
If ipos <> 0 Then bFound = True
ipos = InStr(1, LCase(sTMP), "public property", vbBinaryCompare)
If ipos <> 0 Then bFound = True
ipos = InStr(1, LCase(sTMP), "private property", vbBinaryCompare)
If ipos <> 0 Then bFound = True

ISWRITABLE = bFound
End Function
