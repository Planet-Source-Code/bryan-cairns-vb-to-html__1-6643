Attribute VB_Name = "Module1"
'VB Parseing and Documenting procedures
'Create a new module for C++ Delphi or any other language

'project globals
Global ProjectName As String
Global ProjectVersion As String
Global sLastFile As String
Global Prodir As String
Global ProType As String
Global ProMinVer As String
Global DocDir As String
'3 colors from the RTF File Color Table
Global Hexx1 As String
Global Hexx2 As String
Global Hexx3 As String

Public Function OpenTextFile(sfile As String) As String
On Error GoTo EH
Dim TMPTXT As String
Dim FinTxt As String
Dim IFile As Integer
IFile = FreeFile
Open sfile For Binary Access Read As #IFile
TMPTXT = Space$(LOF(IFile))
Get #IFile, , TMPTXT
Close #IFile
OpenTextFile = TMPTXT
Exit Function
EH:
OpenTextFile = ""
Exit Function
End Function
Public Function CheckFile(sfile As String) As Boolean
On Error Resume Next
Dim Iret
Iret = Dir(sfile)
If Iret > "" Then
CheckFile = True
Else
If Iret = "" Then
CheckFile = False
End If
End If

End Function
Public Function FileLines(sfile As String) As Integer
Dim TextLine
Dim I As Integer
I = 0
Open sfile For Input As #1
Do While Not EOF(1)
   Line Input #1, TextLine
I = I + 1
Loop
Close #1
FileLines = I

End Function
Public Function ParsePath(ByVal TempPath As String, ReturnType As Integer)

    Dim DriveLetter As String
    Dim DirPath As String
    Dim fname As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 And ReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If

        DriveLetter = ""
        DirPath = ""
        fname = ""
        Extension = ""

        If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
            DriveLetter = Left(TempPath, 2)
            TempPath = Mid(TempPath, 3)
        End If

            PathLength = Len(TempPath)

            For Offset = PathLength To 1 Step -1 ' Find the next delimiter.
                Select Case Mid(TempPath, Offset, 1)
                 Case ".": ' This indicates either an extension or a . or a ..
                 ThisLength = Len(TempPath) - Offset

                 If ThisLength >= 1 Then ' Extension
                     Extension = Mid(TempPath, Offset, ThisLength + 1)
                 End If

                     TempPath = Left(TempPath, Offset - 1)
                     Case "\": ' This indicates a path delimiter.
                     ThisLength = Len(TempPath) - Offset

                     If ThisLength >= 1 Then ' Filename
                         fname = Mid(TempPath, Offset + 1, ThisLength)
                         TempPath = Left(TempPath, Offset)
                         FileNameFound = True
                         Exit For
                     End If

                         Case Else
                    End Select

                    Next Offset


                        If FileNameFound = False Then
                            fname = TempPath
                        Else
                            DirPath = TempPath
                        End If


                            If ReturnType = 0 Then
                                ParsePath = DriveLetter
                            ElseIf ReturnType = 1 Then
                                ParsePath = DirPath
                            ElseIf ReturnType = 2 Then
                                ParsePath = fname
                            ElseIf ReturnType = 3 Then
                                ParsePath = Extension
                            End If

End Function
Public Sub CheckTMPDir(sDir As String, dKill As Boolean)
On Error Resume Next
Dim Iret
Iret = Dir(sDir, vbDirectory)
If Iret > "" And dKill = True Then
RmTree sDir
MkDir sDir
Else
If Iret = "" Then
MkDir sDir
End If
End If

End Sub

Public Sub RmTree(ByVal vDir As Variant)
On Error Resume Next
Dim vFile As Variant
    ' Check if "\" was placed at end
    ' If So, Remove it
If Right(vDir, 1) = "\" Then
        vDir = Left(vDir, Len(vDir) - 1)
    End If
' Check if Directory is Valid
    ' If Not, Exit Sub
    vFile = Dir(vDir, vbDirectory)
If vFile = "" Then
        Exit Sub
    End If
' Search For First File
    vFile = Dir(vDir & "\", vbDirectory)
    ' Loop Until All Files and Directories
    ' Have been Deleted
Do Until vFile = ""


        If vFile = "." Or vFile = ".." Then
            vFile = Dir
        ElseIf (GetAttr(vDir & "\" & vFile) And _
            vbDirectory) = vbDirectory Then
            RmTree vDir & "\" & vFile
            vFile = Dir(vDir & "\", vbDirectory)
        Else
            Kill vDir & "\" & vFile
            vFile = Dir
        End If


    Loop


    ' Remove Top Most Directory
    RmDir vDir
End Sub
Public Function Colorss(sLine As String)
Dim S1, S2, S3, r, g, b As String
Dim sTMP As String
Dim ipos As Integer
Dim epos As Integer
Dim TTLine As String
Dim Icount As Integer
Dim Col1 As Long
Dim Col2 As Long
Dim Col3 As Long
Icount = 0
ipos = 0
epos = 1
'The Color Table will look like:
'{\colortbl\red0\green0\blue0;\red0\green0\blue128;\red0\green128\blue0;}
'Add and Parse Hex# variables to get more colors
'from the RTF file color table
'I limited this to 3 colors

sTMP = Mid(sLine, 11, Len(sLine))
'\red0\green0\blue0;\red0\green0\blue128;\red0\green128\blue0;}
'read the line
ipos = InStr(epos, sTMP, ";", vbBinaryCompare)
If ipos = 0 Then Exit Function
TTLine = Mid(sTMP, epos, ipos)
r = GetLinEle(TTLine, "\red", "\green")
g = GetLinEle(TTLine, "\green", "\blue")
b = GetLinEle(TTLine, "\blue", ";")
Col1 = RGB(Int(r), Int(g), Int(b))
Hexx1 = GETHex(Col1)
epos = ipos + 1

ipos = InStr(epos, sTMP, ";", vbBinaryCompare)
If ipos = 0 Then Exit Function
TTLine = Mid(sTMP, epos, ipos)
r = GetLinEle(TTLine, "\red", "\green")
g = GetLinEle(TTLine, "\green", "\blue")
b = GetLinEle(TTLine, "\blue", ";")
Col2 = RGB(Int(r), Int(g), Int(b))
Hexx2 = GETHex(Col2)
epos = ipos + 1

ipos = InStr(epos, sTMP, ";", vbBinaryCompare)
If ipos = 0 Then Exit Function
TTLine = Mid(sTMP, epos, ipos)
r = GetLinEle(TTLine, "\red", "\green")
g = GetLinEle(TTLine, "\green", "\blue")
b = GetLinEle(TTLine, "\blue", ";")
Col3 = RGB(Int(r), Int(g), Int(b))
Hexx3 = GETHex(Col3)
End Function
Public Function GetLinEle(Origin As String, Sep1 As String, Sep2 As String) As String
'Parses a Line of text
On Error GoTo EH
Dim Bpos As Long
Dim epos As Long
Bpos = InStr(1, Origin, Sep1, vbBinaryCompare)
If Bpos = 0 Then Exit Function
epos = InStr(1, Origin, Sep2, vbBinaryCompare)
If Bpos = 0 Then Exit Function
Bpos = Bpos + Len(Sep1)
GetLinEle = Mid(Origin, Bpos, epos - Bpos)
Exit Function
EH:
GetLinEle = ""
Exit Function
End Function
Public Function GETHex(stColor As Long) As String
On Error Resume Next
'stColor = m_CurHex
       '     'If r > 255 Then Exit Sub
       '     'If g > 255 Then Exit Sub
       '     'If b > 255 Then Exit Sub
       Dim r, b, g As Long
       
       Dim dts As Variant
       Dim q, w, e As Variant
       Dim qw, we, gq As Variant
       Dim lCol As Long
       lCol = stColor
       r = lCol Mod &H100
       lCol = lCol \ &H100
       g = lCol Mod &H100
       lCol = lCol \ &H100
       b = lCol Mod &H100
       
       '     'Get Red Hex
       q = Hex(r)

              If Len(q) < 2 Then
                     qw = q
                     q = "0" & qw
              End If

       '     'Get Blue Hex
       w = Hex(b)

              If Len(w) < 2 Then
                     we = w
                     w = "0" & we
              End If

       '     'Get Green Hex
       e = Hex(g)

              If Len(e) < 2 Then
                     gq = e
                     e = "0" & gq
              End If

       'GETRGB = "#" & q & e & w
       GETHex = "#" & q & e & w   '"#" &
End Function
Function RTF2HTML(strRTF As String) As String
    'Version 2.1 (3/30/99)
    'The most current version of this function is available at
    'http://www2.bitstream.net/~bradyh/downl
    '     oads/rtf2html.zip
    'Converts Rich Text encoded text to HTML
Dim ipos As Integer
Dim epos As Integer
Dim ssColTBL As String
ipos = InStr(1, strRTF, "{\colortbl", vbBinaryCompare)
epos = InStr(ipos + 1, strRTF, "}", vbBinaryCompare)
If ipos <> 0 And epos <> 0 Then

ssColTBL = Mid(strRTF, ipos, epos - ipos)
Colorss ssColTBL
Else
If ipos = 0 Or epos = 0 Then
Hexx1 = "#000000"
Hexx2 = "#000000"
Hexx3 = "#000000"
End If
End If
    '     format
        'if you find some text that this function doesn't
        'convert properly please email the text
        '     to
        'bradyh@bitstream.net
        Dim strHTML As String
        Dim l As Long
        Dim lTmp As Long
        Dim lRTFLen As Long
        Dim lBOS As Long 'beginning of section
        Dim lEOS As Long 'end of section
        Dim strTmp As String
        Dim strTmp2 As String
        Dim strEOS 'string To be added to End of section
        Const gHellFrozenOver = False 'always false
        Dim gSkip As Boolean 'skip To Next word/command
        Dim strCodes As String 'codes For ascii To HTML char conversion
        strCodes = "  {00}© {a9}´ {b4}« {ab}» {bb}¡ {a1}¿{bf}À{c0}à{e0}Á{c1}"
        strCodes = strCodes & "á{e1}Â {c2}â {e2}Ã{c3}ã{e3}Ä {c4}ä {e4}Å {c5}å {e5}Æ {c6}"
        strCodes = strCodes & "æ {e6}Ç{c7}ç{e7}Ð{d0}ð{f0}È{c8}è{e8}É{c9}é{e9}Ê {ca}"
        strCodes = strCodes & "ê {ea}Ë {cb}ë {eb}Ì{cc}ì{ec}Í{cd}í{ed}Î {ce}î {ee}Ï {cf}"
        strCodes = strCodes & "ï {ef}Ñ{d1}ñ{f1}Ò{d2}ò{f2}Ó{d3}ó{f3}Ô {d4}ô {f4}Õ{d5}"
        strCodes = strCodes & "õ{f5}Ö {d6}ö {f6}Ø{d8}ø{f8}Ù{d9}ù{f9}Ú{da}ú{fa}Û {db}"
        strCodes = strCodes & "û {fb}Ü {dc}ü {fc}Ý{dd}ý{fd}ÿ {ff}Þ {de}þ {fe}ß {df}§ {a7}"
        strCodes = strCodes & "¶ {b6}µ {b5}¦{a6}±{b1}·{b7}¨{a8}¸ {b8}ª {aa}º {ba}¬{ac}"
        strCodes = strCodes & "­{ad}¯ {af}°{b0}¹ {b9}² {b2}³ {b3}¼{bc}½{bd}¾{be}× {d7}"
        strCodes = strCodes & "÷{f7}¢ {a2}£ {a3}¤{a4}¥{a5}"
        strHTML = ""
        lRTFLen = Len(strRTF)
        'seek first line with text on it
        lBOS = InStr(strRTF, vbCrLf & "\deflang")
        If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 2
        lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
        If lEOS = 0 Then GoTo finally


        While Not gHellFrozenOver
            strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
            l = lBOS


            While l <= lEOS
                strTmp = Mid(strRTF, l, 1)

                Select Case strTmp
                    Case "<"
                    strHTML = strHTML & "&lt;"
                    l = l + 1
                    Case ">"
                   strHTML = strHTML & "&gt;"
                   l = l + 1
                    Case "{"
                    l = l + 1
                    Case "}"
                    strHTML = strHTML & strEOS
                    l = l + 1
                    Case "\" 'special code
                    l = l + 1
                    strTmp = Mid(strRTF, l, 1)
                    '//////////////////////
                    'Below is my modification
                    'to get the colors form the RTF
                    'color table
                    'colors are \cf#  the # is the color table
                    'colors are as follows
                    'cf0 = &h0 //black - #000000
                    'cf1 = &H00800000& //blue - keyword #000080
                    'cf2 =  &H00008000& //green - comment #008000
                    '
                    '
                    Dim bcColor As String
                    Dim CCOlON As Boolean
                    
                    bcColor = Mid(strRTF, l, 3)
                   
                    If bcColor = "pla" And CCOlON = True Then
                        strHTML = strHTML & "</FONT>"
                        CCOlON = False
                    End If
                    If bcColor = "cf1" Then 'color1 - keywords
                        strHTML = strHTML & "<FONT COLOR=""" & Hexx2 & """>"
                        CCOlON = True
                    End If
                    If bcColor = "cf2" Then 'color2 - comments
                        strHTML = strHTML & "<FONT COLOR=""" & Hexx3 & """>"
                        CCOlON = True
                    End If
                    '/////////////////////////////
                    Select Case strTmp
                        Case "b"


                        If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                            strHTML = strHTML & "<B>"
                            strEOS = "</B>" & strEOS
                            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                        ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
                            strHTML = strHTML & "•" 'bullet
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "e"


                        If (Mid(strRTF, l, 7) = "emdash ") Then
                            strHTML = strHTML & "—"
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "i"


                        If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                            strHTML = strHTML & "<I>"
                            strEOS = "</I>" & strEOS
                            If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                        Else
                            gSkip = True
                        End If
                        Case "l"


                        If (Mid(strRTF, l, 10) = "ldblquote ") Then
                            strHTML = strHTML & "“"
                            l = l + 9
                        ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
                            strHTML = strHTML & "‘"
                            l = l + 6
                        Else
                            gSkip = True
                        End If
                        Case "p"


                        If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
                            strHTML = strHTML & strEOS
                            strEOS = ""
                            If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5 'catch Next \ but skip a space
                        Else
                            gSkip = True
                        End If
                        Case "r"


                        If (Mid(strRTF, l, 7) = "rquote ") Then
                            strHTML = strHTML & "’"
                            l = l + 6
                        ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
                            strHTML = strHTML & "”"
                            l = l + 9
                        Else
                            gSkip = True
                        End If
                        Case "t"


                        If (Mid(strRTF, l, 4) = "tab ") Then
                            strHTML = strHTML & Chr$(9) 'tab
                            l = l + 3
                        Else
                            gSkip = True
                        End If
                        Case "'"
                        strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
                        lTmp = InStr(strCodes, strTmp2)


                        If lTmp = 0 Then
                            strHTML = strHTML & Chr("&H" & Mid(strTmp2, 2, 2))
                        Else
                            strHTML = strHTML & Trim(Mid(strCodes, lTmp - 8, 8))
                        End If
                        l = l + 2
                        Case "~"
                        strHTML = strHTML & " "
                        Case "{", "}", "\"
                        strHTML = strHTML & strTmp
                        Case vbLf, vbCr, vbCrLf 'always use vbCrLf
                        strHTML = strHTML & vbCrLf
                        Case Else
                        gSkip = True
                    End Select


                If gSkip = True Then
                    'skip everything up until the next space
                    '     or "\"


                    While ((Mid(strRTF, l, 1) <> " ") And (Mid(strRTF, l, 1) <> "\"))
                        l = l + 1
                    Wend
                    gSkip = False
                    If (Mid(strRTF, l, 1) = "\") Then l = l - 1
                End If
                l = l + 1
                Case vbLf, vbCr, vbCrLf
                l = l + 1
                Case Else
                strHTML = strHTML & strTmp
                l = l + 1
            End Select
    Wend
    lBOS = lEOS + 2
    lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally
    strHTML = strHTML & "<br>"
Wend
finally:
RTF2HTML = strHTML
End Function
