Attribute VB_Name = "Coloring"
'This is the Module that Colorizes the contents of a
'richtextbox  - defined by keywords
'I did not write or modify this so if you use this
'please give VBDiamond the credit
Private gsBlackKeywords    As String
Private gsBlueKeyWords     As String

Public Sub ColorizeWords(rtf As RichTextBox)
    'VBDiamond
   ' * Web Site     : www.geocities.com/ResearchTriangle/6311/
   
   Dim sBuffer    As String
   Dim nI         As Long
   Dim nJ         As Long
   Dim sTmpWord   As String
   Dim nStartPos  As Long
   Dim nSelLen    As Long
   Dim nWordPos   As Long
   
   'Dim cHourglass    As class_Hourglass
   'Set cHourglass = New class_Hourglass
   
   sBuffer = rtf.Text
   sTmpWord = ""
   With rtf
      For nI = 1 To Len(sBuffer)
         Select Case Mid(sBuffer, nI, 1)
        Case "A" To "Z", "a" To "z", "_"
           If sTmpWord = "" Then nStartPos = nI
           sTmpWord = sTmpWord & Mid(sBuffer, nI, 1)
        
        Case Chr(34)
           nSelLen = 1
           For nJ = 1 To 9999999
              If Mid(sBuffer, nI + 1, 1) = Chr(34) Then
             nI = nI + 2
             Exit For
              Else
             nSelLen = nSelLen + 1
             nI = nI + 1
              End If
           Next
        
        Case Chr(39)
           .SelStart = nI - 1
           nSelLen = 0
           For nJ = 1 To 9999999
              If Mid(sBuffer, nI, 2) = vbCrLf Then
             Exit For
              Else
             nSelLen = nSelLen + 1
             nI = nI + 1
              End If
           Next
           .SelLength = nSelLen
           .SelColor = RGB(0, 127, 0)
        
        Case Else
           If Not (Len(sTmpWord) = 0) Then
              .SelStart = nStartPos - 1
              .SelLength = Len(sTmpWord)
              nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
              If nWordPos <> 0 Then
             .SelColor = RGB(0, 0, 0)
             .SelText = Mid(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
              End If
              nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
              If nWordPos <> 0 Then
             .SelColor = RGB(0, 0, 127)
             .SelText = Mid(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
              End If
              If UCase(sTmpWord) = "REM" Then
             .SelStart = nI - 4
             .SelLength = 3
             For nJ = 1 To 9999999
                If Mid(sBuffer, nI, 2) = vbCrLf Then
                   Exit For
                Else
                   .SelLength = .SelLength + 1
                   nI = nI + 1
                End If
             Next
             .SelColor = RGB(0, 127, 0)
             .SelText = LCase(.SelText)
              End If
           End If
           sTmpWord = ""
         End Select
      Next
      .SelStart = 0
   
   End With
   
End Sub

Public Sub InitColorize()
   
   gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
   gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*"

End Sub
