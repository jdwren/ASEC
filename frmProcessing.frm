VERSION 5.00
Begin VB.Form frmProcessing 
   ClientHeight    =   6150
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   10575
   End
   Begin VB.CommandButton btnORerrors 
      Caption         =   "Extract ratio-confidence interval patterns from MEDLINE records"
      Height          =   555
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   4575
   End
   Begin VB.CommandButton btnDivErr 
      Caption         =   "Extract percent-division patterns from MEDLINE records"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton btnConvert 
      Caption         =   "Convert MEDLINE XML files to simpler, more compact format"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label lblProgress 
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   3120
      Width           =   5895
   End
End
Attribute VB_Name = "frmProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConvert_Click()
Dim a$, b$, C$, FileName$, FileOut$, dbname$, fstr$, acro$, def$, yr$, mon$, da$, tit$, lcb$, newstring$, newdbname$, AssayMatches$, ch As String * 1
Dim sp$(), sp2$()
Dim bytesread&, fsize&, BlockSize&, t&, u&, v&, f&, sentencestart&, Lparen&, Rparen&, UB&, UB2&
Dim abstr&, AbStart&, abend&, recs&, TotalRecs&, Dfield&, textlen&, acro_count&, DupePMIDs&
Dim prog!
Dim PubDatePos&, lastPDpos&, nextPDpos&, nextabpos&, PMIDpos&, PMIDend&, PMID&, NewObjsToAdd&, CoPos&, CoPos2&, ATpos&, ATpos2&, TitStart&, TitEnd&, FirstPos&, NextPos&
Dim FileNum1%, FileNum2%
Dim endflag%, loopflag%, flag%, AddNewObjsFlag%, NOAflag%
Dim HistoryFlag%, TermExpansionFlag%, CoOccurrenceDumpFlag%, AssayFlag%, DumpFlag%, APflag%, AbstractFoundflag%
Dim fs As Variant

Set fs = CreateObject("Scripting.FileSystemObject")
inp$ = vbNullString
frmProcessing.MousePointer = vbHourglass
TotalAbs = 0
BlockSize = 60000
For f = 1 To 1259
    t = 0: u = 0
    If f < 10 Then
        fstr = "000" & Trim$(Str(f))
    ElseIf f < 100 Then
        fstr = "00" & Trim$(Str(f))
    ElseIf f < 1000 Then
        fstr = "0" & Trim$(Str(f))
    Else
        fstr = Trim$(Str(f))
    End If
    FileName$ = "D:\MEDLINE\medline17n" & fstr & ".xml"             'point this to the directory with MEDLINE XML files
    FileOut$ = "D:\MEDLINE processed\medline17n" & fstr & ".xml"    'This will be where the records in summary form are put
    lblProgress.Caption = "Processing File# " & fstr: lblProgress.Refresh
    fsize = FileLen(FileName$)
    FileNum1 = FreeFile
    Open FileName$ For Input As FileNum1
    FileNum2 = FreeFile
    Open FileOut$ For Output As FileNum2
    bytesread = 0: recs = 0: abstr = 0
   
    While Not EOF(FileNum1)
        'Grab a chunk of text representing a set of complete articles
        If bytesread + BlockSize > fsize Then
            a$ = Input(fsize - bytesread, FileNum1)
            bytesread = bytesread + Len(a$)
        Else
            a$ = Input(BlockSize, FileNum1)
            bytesread = bytesread + Len(a$)
            b$ = vbNullString: C$ = vbNullString: endflag = 0
            While endflag = 0                   'Extend block read to end of article (to avoid cutting key fields in half)
                b$ = b$ & Input(1, FileNum1)
                bytesread = bytesread + 1
                If bytesread = fsize Then endflag = 1
                If Right$(b$, 18) = "</MedlineCitation>" Then endflag = 1
                If Len(b$) > 60000 Then         'this step just to improve speed
                    C$ = C$ & b$: b$ = vbNullString
                End If
            Wend
            a$ = a$ & C$ & b$
        End If
        
        'PROCESS TEXT BLOCK
        sp$ = Split(a$, "</MedlineCitation>")                           'split the block by citation record
        UB = UBound(sp$)
        For t = 0 To UB - 1
            recs = recs + 1: TotalRecs = TotalRecs + 1
            dat$ = "99999999": PMID = 0: JName$ = vbNullString
            
            CoPos = InStr(1, sp$(t), "<CommentsCorrectionsList>")       'This is where the citations start. Cut it out so no PMID confusion
            If CoPos > 0 Then sp$(t) = Left(sp$(t), CoPos)
            'get PMID
            PMIDpos = InStr(1, sp$(t), "<PMID Version=")
            If PMIDpos > 0 Then
                PMIDend = InStr(PMIDpos, sp$(t), ">")
                PMID = Val(Mid$(sp$(t), PMIDend + 1, InStr(PMIDend, sp$(t), "</PMID>") - PMIDpos - 1))
            End If
            
            'get Publication Date
            PubDatePos = InStr(1, sp$(t), "<ArticleDate")   'Try article date and, if not there, try pubdate
            If PubDatePos = 0 Then
                PubDatePos = InStr(1, sp$(t), "<PubDate") 'Get date start
                lastPDpos = InStr(1, sp$(t), "</PubDate>") 'Get date end
            Else
                lastPDpos = InStr(1, sp$(t), "</ArticleDate>") 'Get date end
            End If
            If PubDatePos > 0 Then
                Dfield = InStr(PubDatePos, sp$(t), "<Year>")
                If Dfield > 0 And Dfield < lastPDpos Then yr$ = Mid$(sp$(t), Dfield + 6, 4)
                Dfield = InStr(PubDatePos, sp$(t), "<Month>")
                If Dfield > 0 And Dfield < lastPDpos Then mon$ = Mid$(sp$(t), Dfield + 7, 3)
                Call Convert_Month(mon$)
                If mon$ = "unknown" Then mon$ = "01" ' Default to January
                Dfield = InStr(PubDatePos, sp$(t), "<Day>")
                If Dfield > 0 And Dfield < lastPDpos Then
                    da$ = Mid$(sp$(t), Dfield + 5, 2)
                    If Right$(da$, 1) = "<" Then da$ = "0" & Left$(da$, 1)
                Else
                    da$ = "00"
                End If
                dat$ = yr$ & mon$ & da$
            End If
            
            'Get journal name
            JName$ = vbNullString
            JNstart = InStr(1, sp$(t), "<MedlineTA>")                 'Medline journal title abbrev
            If JNstart > 0 Then
                JNstart = JNstart + 11
                JNend = InStr(JNstart, sp$(t), "</MedlineTA>")
                JName$ = Mid$(sp$(t), JNstart, JNend - JNstart)
                JName$ = Trim(JName$)
            End If

            'Get title
            TitStart = InStr(1, sp$(t), "<ArticleTitle>")
            TitEnd = InStr(1, sp$(t), "</ArticleTitle>")
            If TitStart > 0 And TitEnd > 0 Then
                tit$ = Mid$(sp$(t), TitStart + 14, TitEnd - TitStart - 14)
                If Left$(tit$, 1) = "[" Then
                    If Right$(tit$, 2) = "]." Then
                        tit$ = Mid$(tit$, 2, Len(tit$) - 3) & "."             'remove marks indicating article was translated
                    End If
                End If
                b$ = " " & tit$
            Else
                TitStart = 0
            End If
            'Get abstract
            AbstractFoundflag = 0
            AbStart = InStr(1, sp$(t), "<Abstract>")
            If AbStart > 0 Then
                AbstractFoundflag = 1
                abstr = abstr + 1: TotalAbs = TotalAbs + 1
                AbStart = InStr(AbStart + 1, sp$(t), ">") + 1   'look for closing >
                abend = InStr(AbStart, sp$(t), "</Abstract>")
                If abend > 0 Then
                    b$ = b$ & " " & Mid$(sp$(t), AbStart, abend - AbStart) & " "
                End If
                If InStr(b$, "</AbstractText>") > 0 Then
                    b$ = Replace(b$, "</AbstractText>", " ")
                End If
                While InStr(b$, "<AbstractText") > 0            'Concatenate subcategories (if it has them)
                    ATpos = InStr(b$, "<AbstractText")
                    ATpos2 = InStr(ATpos + 10, b$, ">")
                    If ATpos > 0 Then           '> string starting pos
                        b$ = Left$(b$, ATpos - 1) & Right$(b$, Len(b$) - ATpos2)
                    Else
                        b$ = Right$(b$, Len(b$) - ATpos2)
                    End If
                Wend
                'Start streamlining the text
                If InStr(b$, vbCr) > 0 Then b$ = Replace(b$, vbCr, " ")
                If InStr(b$, vbLf) > 0 Then b$ = Replace(b$, vbLf, " ")
                If InStr(b$, Chr$(34)) > 0 Then b$ = Replace(b$, Chr$(34), vbNullString)
                While InStr(b$, "  ") > 0       'eliminate double spaces
                    b$ = Replace(b$, "  ", " ")
                Wend
                If InStr(b$, "&lt;") > 0 Then b$ = Replace(b$, "&lt;", "<")
                If InStr(b$, "&gt;") > 0 Then b$ = Replace(b$, "&gt;", ">")
                If InStr(b$, "&amp;") > 0 Then b$ = Replace(b$, "&amp;", "&")
                If InStr(b$, "&quot;") > 0 Then b$ = Replace(b$, "&quot;", "`")
                If InStr(b$, "&mdash;") > 0 Then b$ = Replace(b$, "&mdash;", "-")
                If InStr(b$, "�") > 0 Then b$ = Replace(b$, "�", ".")
                If InStr(b$, "∼") > 0 Then b$ = Replace(b$, "∼", "-")
                If InStr(b$, "‒") > 0 Then b$ = Replace(b$, "‒", "-")
                If InStr(b$, "–") > 0 Then b$ = Replace(b$, "–", "-")
                If InStr(b$, "≤") > 0 Then b$ = Replace(b$, "≤", "<=")
                If InStr(b$, "≥") > 0 Then b$ = Replace(b$, "≥", ">=")
                If InStr(b$, "±") > 0 Then b$ = Replace(b$, "±", "+/-")
                If InStr(b$, "×") > 0 Then b$ = Replace(b$, "×", "x")     'multiplication symbol
                If InStr(b$, " ") > 0 Then b$ = Replace(b$, " ", " ")
                If InStr(b$, "( ") > 0 Then b$ = Replace(b$, "( ", "(")
                If InStr(b$, " )") > 0 Then b$ = Replace(b$, " )", ")")
                If InStr(b$, " per cent ") > 0 Then b$ = Replace(b$, " per cent ", " percent ")
                If InStr(b$, "less than or equal to") > 0 Then b$ = Replace(b$, "less than or equal to", "<=")
                If InStr(b$, "p less than ") > 0 Then b$ = Replace(b$, "p less than ", "p<")
                If InStr(b$, "P less than ") > 0 Then b$ = Replace(b$, "P less than ", "p<")
                If InStr(b$, " confidence interval [CI]") > 0 Then b$ = Replace(b$, "95% confidence interval [CI]", " CI")
                If InStr(b$, " confidence interval (CI)") > 0 Then b$ = Replace(b$, "95% confidence interval (CI)", " CI")
                If InStr(b$, " confidence interval") > 0 Then b$ = Replace(b$, "95% confidence interval", " CI")
                If InStr(b$, "odds ratio [OR]=") > 0 Then b$ = Replace(b$, "odds ratio [OR]=", "OR=")
                If InStr(b$, "odds ratio (OR)=") > 0 Then b$ = Replace(b$, "odds ratio (OR)=", "OR=")
                If InStr(b$, "< or =") > 0 Then b$ = Replace(b$, "< or =", "<=")
                If InStr(b$, "> or =") > 0 Then b$ = Replace(b$, "> or =", ">=")
                If InStr(b$, " +/- ") > 0 Then b$ = Replace(b$, " +/- ", "+/-")
                If InStr(b$, " <= ") > 0 Then b$ = Replace(b$, " <= ", "<=")
                If InStr(b$, " = ") > 0 Then b$ = Replace(b$, " = ", "=")
                If InStr(b$, " < ") > 0 Then b$ = Replace(b$, " < ", "<")
                If InStr(b$, " > ") > 0 Then b$ = Replace(b$, " > ", ">")
                If InStr(b$, ", ,") > 0 Then b$ = Replace(b$, ", ,", ",")
                If InStr(b$, " vs. ") > 0 Then b$ = Replace(b$, " vs. ", " vs ")
                If InStr(b$, " v. ") > 0 Then b$ = Replace(b$, " v. ", " vs ")
            End If
                
            If (TitStart > 0 Or AbStart > 0) And APflag = 0 Then        'only output those with at least a title & abstract
                'just output key fields
                Print #FileNum2, "<PubDate>" & dat$ & "</PubDate>"
                Print #FileNum2, "<PMID>" & PMID & "</PMID>"
                Print #FileNum2, "<Journal>" & JName$ & "</Journal>"
                Print #FileNum2, "<RecordText>" & b$ & "</RecordText>"
            End If
        Next t
        inp$ = " ": b$ = vbNullString
    Wend
lblProgress.Caption = vbNullString
Close FileNum1
Close FileNum2
Debug.Print "File# " & f & "   Recs: " & recs & "   TotalRecs:" & TotalRecs
recs = 0
Next f

frmProcessing.MousePointer = vbDefault
End Sub

Private Sub btnDivErr_Click()
Dim a$, b$, C$, E$, O$, FileName$, FileOut$, dbname$, fstr$, yr$, mon$, da$, tit$, lcb$, ch As String * 1, repl$, ReportedNum$, ProblemFlag$
Dim LastWord$, TwoLastWords$
Dim ValidPct$, ValidLeft$, ValidRight$, ValidRatio$, Context$, LCContext$, Ratio$, Percent$, Percent2$, LeftNum$, RightNum$, Rightnum2$, Leftnum2$
Dim sp$(), sp2$(), RecSP$(), WordCheck%()
Dim dv@, n@, d@, pct@, pct2@, diff@, dv2@, diff2@, log10diff@, CalcNum@
Dim bytesread&, fsize&, BlockSize&, t&, u&, v&, f&, sentencestart&, Lparen&, Rparen&, UB&, UB2&, RecCt&, LastPeriod&
Dim abstr&, AbStart&, abend&, recs&, TotalRecs&, Dfield&, textlen&, acro_count&
Dim PubDatePos&, lastPDpos&, nextPDpos&, nextabpos&, PMIDpos&, PMID&, NewObjsToAdd&, CoPos&, CoPos2&, ATpos&, ATpos2&, TitStart&, FirstPos&, NextPos&
Dim FileNum1%, FileNum2%
Dim endflag%, loopflag%, flag%, ExamineFlag%, OutputFlag%
Dim fs As Variant

Set fs = CreateObject("Scripting.FileSystemObject")
inp$ = vbNullString
frmProcessing.MousePointer = vbHourglass
TotalAbs = 0
BlockSize = 60000

FileNum2 = FreeFile
FileOut$ = "D:\ASEC MEDLINE division errors.txt"
Open FileOut$ For Output As FileNum2
Print #FileNum2, "PMID" & vbTab & "rep. ratio" & vbTab & "rep. pct" & vbTab & "calc pct" & vbTab & "diff" & vbTab & "log10diff" & vbTab & "flags" & vbTab & "Flags" & vbTab & "context"

For f = 144 To 1259
    If f < 10 Then
        fstr = "000" & Trim$(Str(f))
    ElseIf f < 100 Then
        fstr = "00" & Trim$(Str(f))
    ElseIf f < 1000 Then
        fstr = "0" & Trim$(Str(f))
    Else
        fstr = Trim$(Str(f))
    End If
    FileName$ = "D:\MEDLINE Processed\medline17n" & fstr & ".xml"
    lblProgress.Caption = "Processing File# " & fstr: lblProgress.Refresh
    fsize = FileLen(FileName$)
    FileNum1 = FreeFile
    Open FileName$ For Input As FileNum1
    bytesread = 0: recs = 0: abstr = 0
   
    While Not EOF(FileNum1)
        'Grab a chunk of text representing a set of complete articles
        If bytesread + BlockSize > fsize Then
            a$ = Input(fsize - bytesread, FileNum1)
            bytesread = bytesread + Len(a$)
        Else
            a$ = Input(BlockSize, FileNum1)
            bytesread = bytesread + Len(a$)
            b$ = vbNullString: endflag = 0
            While endflag = 0           'Extend block read to end of article (to avoid cutting key fields in half)
                b$ = b$ & Input(1, FileNum1)
                bytesread = bytesread + 1
                If bytesread = fsize Then endflag = 1
                If Right$(b$, 13) = "</RecordText>" Then endflag = 1
            Wend
            a$ = a$ & b$
        End If
        
        'PROCESS TEXT BLOCK
        endflag = 0: lastPDpos = 1
        RecSP$ = Split(a$, "</RecordText>")
        RecUB = UBound(RecSP$) - 1
        For RecCt = 0 To RecUB
            PubDatePos = InStr(RecSP$(RecCt), "</PubDate>")
            dat$ = Mid$(RecSP$(RecCt), PubDatePos - 8, 8)
            PMIDst = InStr(RecSP$(RecCt), "<PMID>")
            PMIDend = InStr(RecSP$(RecCt), "</PMID>")
            PMID = Val(Mid$(RecSP$(RecCt), PMIDst + 6, PMIDend - PMIDst - 6))
            AbStart = InStr(RecSP$(RecCt), "<RecordText>")
            b$ = " " & Right$(RecSP$(RecCt), Len(RecSP$(RecCt)) - AbStart - 12) & " "
            LastPeriod = 0
            
            If PMID = 12859115 Or PMID = 12923312 Then AbStart = 0       'problematic PMIDs
            If AbStart > 0 Then
                sp$ = Split(b$, " ")
                UB = UBound(sp$)
                Erase WordCheck
                ReDim WordCheck%(UB)
                For t = 0 To UB - 1
                    If Right$(sp$(t), 1) = "." Then LastPeriod = t
                    Ratio$ = "": Percent$ = "": ExamineFlag = 0
                    'look for "X/Y (Z%)" pattern
                    If InStr(sp$(t), "/") > 0 Then               'potential divisor sign
                    If Not Right$(sp$(t), 1) Like "[.,;:]" Then    'the two values aren't split by the end of a sentence or phrase
                    If InStr(sp$(t), "+/-") = 0 Then             'make sure the divisor isn't part of an error estimate
                        If InStr(sp$(t + 1), "%)") > 0 Then      'followed by potential percent value reported
                            Ratio$ = sp$(t): Percent$ = sp$(t + 1)
                            ExamineFlag = 1
                        End If
                    End If
                    End If
                    End If
                    'look for "Z% (X/Y)" pattern
                    If InStr(sp$(t), "%") > 0 Then              'percent indicator
                        If InStr(sp$(t + 1), "/") > 0 Then      'next word a potential ratio
                        If Not Right$(sp$(t), 1) Like "[.,;:]" Then    'the two values aren't split by the end of a sentence or phrase
                        If InStr(sp$(t + 1), "+/-") = 0 Then    'make sure it's not an error estimate
                            If (Len(sp$(t + 2)) = 3 And Left$(sp$(t + 2), 3) Like "###") Or Left$(sp$(t + 2), 4) Like "###)" Or Left$(sp$(t + 2), 5) Like "###.#" Or Left$(sp$(t + 2), 4) Like "###[,;:]" Then   'like ###, or ###) or ###; etc
                                sp$(t + 1) = sp$(t + 1) & Left$(sp$(t + 2), 3) 'Europeans use spaces for commas sometimes (e.g., 17 452)
                                sp$(t + 2) = ""
                            End If
                            Ratio$ = sp$(t + 1): Percent$ = sp$(t)
                            ExamineFlag = 1
                        End If
                        End If
                        End If
                    End If
                    If WordCheck(t) = 1 Then ExamineFlag = 0    'This % or # was part of a valid pattern earlier
                    
                    'Now check if these are really ratio & percent value being reported
                    If ExamineFlag = 1 Then
                        Context = vbNullString
                        If LastPeriod > 0 Then
                            u = LastPeriod + 1
                        Else
                            u = 0
                        End If
                        If InStr(sp$(u), "%") > 0 And u >= 4 Then       'the % is the first word of the sentence
                            Context = sp$(u - 4) & " " & sp$(u - 3) & " " & sp$(u - 2) & " " & sp$(u - 1) & " "
                        End If
                        loopflag = 0
                        While loopflag = 0          'get surrounding context
                            Context = Context & sp$(u) & " "
                            If u < UB Then
                                If Right$(sp$(u), 1) = "." Then loopflag = 1
                            Else
                                loopflag = 1
                            End If
                            u = u + 1
                        Wend

                        'extract context for the error
                        Call Format_String(Percent$)
                        Call Is_it_really_a_pct(Percent$, ValidPct$)
                        If ValidPct = "Y" And InStr(Percent$, ",") > 0 Then Percent$ = Replace(Percent$, ",", ".")  'Europeans use commas for decimals (sigh)
                        If ValidPct = "Y" Then
                            Call Format_String(Ratio$)
                            sp2$ = Split(Ratio$, "/")
                            LeftNum$ = sp2$(0)
                            RightNum$ = sp2$(1)
                            Call is_it_really_a_number(LeftNum$, ValidLeft$)
                            Call is_it_really_a_number(RightNum$, ValidRight$)
                            ValidRatio = "Y"
                            If UBound(sp2$) > 1 Then ValidRatio = "N"   'multiple slashes
                            If ValidLeft$ = "N" Or ValidRight$ = "N" Then ValidRatio = "N"
                            If ValidRatio = "Y" Then        'the #s are ok. One more check for year-based patterns
                                If Val(LeftNum$) >= 1900 And Val(LeftNum$) <= 2016 Then
                                    O$ = Right$(LeftNum$, 2)
                                    If Val(RightNum$) = Val(O$) + 1 Then ValidRatio = "N"             'it's a year (e.g., 1998/99)
                                    O$ = Right$(LeftNum$, 1)
                                    If Val(RightNum$) = Val(O$) + 1 Then ValidRatio = "N"             'it's a year (e.g., 1998/9)
                                    If Val(LeftNum$) / Val(RightNum$) > 0.995 Then ValidRatio = "N"    'The two are very close together, suggesting year
                                End If
                            End If
                            If ValidRatio = "Y" Then
                                If Val(RightNum$) <> 0 Then
                                    WordCheck(t) = 1: WordCheck(t + 1) = 1
                                    dv = LeftNum$ / RightNum$      'oddly, using Val() truncates commas in numbers, but string division works
                                    Percent = Replace(Percent$, "%", vbNullString)
                                    pct = Val(Percent$) / 100
                                    If pct <> 0 Then
                                        diff = dv - pct
                                        'Some errors are based on faulty assumptions about syntax
                                        'so examine some alternate possibilities to see if they are reasonable
                                        If diff > 0.1 Then
                                            If InStr(RightNum$, ".") > 0 Then           'Sometimes a decimal is supposed to be a comma
                                                Rightnum2$ = Replace(RightNum$, ".", vbNullString)
                                                dv2 = LeftNum$ / Rightnum2$
                                                diff2 = dv2 - pct
                                                If diff2 < 0.1 And Abs(diff2) < Abs(diff) Then
                                                    dv = dv2: diff = diff2: Ratio = LeftNum$ & "/" & Rightnum2$
                                                End If
                                            End If
                                        End If
                                        If diff > 0.02 Then   'sometimes a space is accidentally inserted
                                            If InStr(sp$(t), "%") > 0 And Right$(sp$(t - 1), 2) Like "#." And IsNumeric(sp$(t - 1)) Then
                                                Percent2$ = sp$(t - 1) & Percent$
                                                pct2 = Val(Percent2$) / 100
                                                diff2 = dv - pct2
                                                If diff2 < 0.02 And Abs(diff2) < Abs(diff) Then
                                                    diff = diff2: Percent$ = Percent2$: pct = pct2
                                                End If
                                            End If
                                        End If
                                        If diff > 0.1 Then
                                            'in instances where they are reporting something that should add up to 100%,
                                            'the numerator/denominator are swapped (e.g., in sensitivity/specificity). Try flipping them and
                                            'see if they are within 1%. If so, swap them and mark with "@"
                                            If Val(LeftNum$) > 0 Then
                                                dv2 = RightNum$ / LeftNum$
                                                diff2 = dv2 - pct
                                                If diff2 < 0.01 And Abs(diff2) < Abs(diff) Then
                                                    dv = dv2: diff = diff2: Ratio = RightNum$ & "/" & LeftNum$ & "@"
                                                End If
                                            End If
                                        End If
                                        
                                        'See if rounding up or rounding down gets closer to the reported number (benefit of doubt)
                                        ReportedNum = Str$(pct)
                                        CalcNum = dv
                                        If CalcNum <> 0 Then Call Choose_round_up_or_down(ReportedNum$, CalcNum@)
                                        diff = CalcNum - pct
                                        
                                        If CalcNum <> 0 And pct <> 0 Then
                                            log10diff = (Log(CalcNum * 100) / Log(10)) - (Log(pct * 100) / Log(10))
                                        Else
                                            log10diff = 0
                                        End If
                                        log10diff = Int(log10diff * 1000) / 1000
                                        
                                        OutputFlag = 1
                                        ProblemFlag = vbNullString: LastWord = vbNullString: TwoLastWords = vbNullString
                                        'screen out words known not to be ratios based on preceding words
                                        If t > 0 Then   'one word
                                            LastWord$ = LCase(sp$(t - 1))
                                            If LastWord$ = "grade" Then OutputFlag = 0
                                            If LastWord$ = "hpv" Then OutputFlag = 0
                                            If LastWord$ = "over" Then ProblemFlag = ProblemFlag & "<>"
                                            If LastWord$ = "almost" Then ProblemFlag = ProblemFlag & "<>"
                                        End If
                                        If t > 1 Then   'two words
                                            TwoLastWords = LCase(sp$(t - 2)) & " " & LCase(sp$(t - 1))
                                            If InStr(TwoLastWords, "ribotype") > 0 Then OutputFlag = 0
                                            If TwoLastWords = "more than" Then ProblemFlag = ProblemFlag & "<>"
                                            If TwoLastWords = "greater than" Then ProblemFlag = ProblemFlag & "<>"
                                            If TwoLastWords = "less than" Then ProblemFlag = ProblemFlag & "<>"
                                        End If
                                            
                                        'flag potential problems based on problematic keywords that are not already screened out
                                        LCContext = LCase(Context$)
                                        If OutputFlag = 1 Then
                                            If InStr(LCContext$, "genotype") > 0 Then ProblemFlag = ProblemFlag & "genotype"
                                            If InStr(LCContext$, "allele") > 0 Then ProblemFlag = ProblemFlag & "allele"
                                            If InStr(LCContext$, "visual acuity") > 0 Then ProblemFlag = ProblemFlag & "vision"
                                            If InStr(LCContext$, "ribotype") > 0 Then ProblemFlag = ProblemFlag & "ribotype"
                                            If InStr(LCContext$, "grade") > 0 Then ProblemFlag = ProblemFlag & "grade"
                                            If InStr(LCContext$, "hpv") > 0 Then ProblemFlag = ProblemFlag & "hpv"
                                        End If
                                        
                                        'screen out known problems by context keywords
                                        If InStr(Context$, "HES") > 0 Or InStr(Context$, "ethyl starch") > 0 Then OutputFlag = 0
                                        If InStr(Context$, "ompound 48/80") > 0 Or InStr(Context$, "crude protein content") > 0 Then OutputFlag = 0
                                        
                                        If OutputFlag = 1 Then
                                            Print #FileNum2, PMID & vbTab & "'" & Ratio$ & vbTab & pct & vbTab & CalcNum & vbTab & diff & vbTab & log10diff & vbTab & ProblemFlag & vbTab & Context
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next t
            End If      'if abstart > 0
        Next RecCt
        inp$ = " ": b$ = vbNullString
    Wend
lblProgress.Caption = vbNullString
Close FileNum1
Next f
Close FileNum2
frmProcessing.MousePointer = vbDefault
End Sub
Public Sub Format_String(ByRef a$)
    If Right$(a$, 1) = "." Then a$ = Left$(a$, Len(a$) - 1)     'only remove periods if at the end of the string
    If Right$(a$, 1) = ";" Then a$ = Left$(a$, Len(a$) - 1)     'remove dividing sems
    If Right$(a$, 1) = "," Then a$ = Left$(a$, Len(a$) - 1)     'remove trailing commas
    If Right$(a$, 1) = ")" Then a$ = Left$(a$, Len(a$) - 1)     'remove boundary parens
    If Left$(a$, 1) = "(" Then a$ = Right$(a$, Len(a$) - 1)     'remove boundary parens
End Sub

Public Sub Is_it_really_a_pct(ByVal a$, ByRef validflag$)
Dim txtlen%, t%
validflag$ = "Y"                    'Should start numeric, end numeric and have only . or , in the middle
a$ = Replace(a$, "%", vbNullString)
If Not Left$(a$, 1) Like "[0-9]" And Left$(a$, 1) <> "-" Then validflag = "N"       'first character numeric or sign
If Not Right$(a$, 1) Like "[0-9]" Then validflag = "N"                              'last character numeric
txtlen = Len(a$)
For t = 1 To txtlen
    If Not Mid$(a$, t, 1) Like "[0-9]" Then
        If Mid$(a$, t, 1) <> "." And Mid$(a$, t, 1) <> "," Then validflag = "N"
    End If
Next t
End Sub

Public Sub is_it_really_a_number(ByVal a$, ByRef validflag$)
Dim txtlen%, t%, pcount%
validflag$ = "Y"                    'Should start numeric, end numeric and have only . or , in the middle
If Not Left$(a$, 1) Like "[0-9]" And Left$(a$, 1) <> "-" Then validflag = "N"       'first character numeric or sign
If Not Right$(a$, 1) Like "[0-9]" Then validflag = "N"                              'last character numeric
txtlen = Len(a$)
For t = 1 To txtlen
    If Not Mid$(a$, t, 1) Like "[0-9]" Then
        If Mid$(a$, t, 1) <> "." And Mid$(a$, t, 1) <> "," Then validflag = "N"
    End If
    If Mid$(a$, t, 1) = "." Then pcount = pcount + 1
Next t
If pcount > 1 Then validflag = "N"
End Sub

Private Sub btnORerrors_Click()
Dim a$, b$, C$, d$, O$, FileName$, FileOut$, dbname$, fstr$, yr$, mon$, da$, tit$, LOR$, lcor$, ch As String * 1, repl$, pval$, ErrFlag$, OR_RR$, OR_RR_lcase$, JName$
Dim ValidLeft$, ValidRight$, Context$, ORstring$, modORstring$, OddsR$, CIlower$, CIupper$, LeftNum$, RightNum$, ReportedNum$, LowestOR$, HighestOR$, NewWord$
Dim UnmodAbstr$
Dim sp$(), sp2$(), sp3$(), RecSP$(), UnmodSP$()
Dim diff@, logoddsR@, logoddsU@, logoddsL@, log10diff@, Middle@, CalcNum@, OddsRdiff@, MinDiff@
Dim bytesread&, fsize&, BlockSize&, t&, u&, v&, f&, sentencestart&, Lparen&, Rparen&, UB&, UB2&, RecCt&, LastPeriod&, sigfigs%
Dim abstr&, AbStart&, abend&, recs&, TotalRecs&, Dfield&, textlen&, txtpos&, JnameSt&, JnameEnd&, lenp&
Dim PubDatePos&, lastPDpos&, nextPDpos&, nextabpos&, PMIDpos&, PMID&, TitStart&, FirstPos&, NextPos&
Dim FileNum1%, FileNum2%
Dim endflag%, loopflag%, flag%, ExamineFlag%, ParenFlag%, ProcessedFlag%
Dim ORflag%, RRflag%, HRflag%
Dim fs As Variant

Set fs = CreateObject("Scripting.FileSystemObject")
inp$ = vbNullString
frmProcessing.MousePointer = vbHourglass
TotalAbs = 0
BlockSize = 60000

FileNum2 = FreeFile
FileOut$ = "D:\ASEC MEDLINE Ratio-CI errors.txt"
Open FileOut$ For Output As FileNum2
Print #FileNum2, "PMID" & vbTab & "Ratio stmt" & vbTab & "R modified" & vbTab & "p-val" & vbTab & "rep. lower" & vbTab & "rep. R" & vbTab & "rep. upper" & vbTab & "calc R" & vbTab & "min R" & vbTab & "exact R" & vbTab & "max R" & vbTab & "min diff" & vbTab & "flag" & vbTab & "Journal" & vbTab & "context"

For f = 214 To 1259          'Process MEDLINE XML files (patterns are sparse before file #214)
    If f < 10 Then
        fstr = "000" & Trim$(Str(f))
    ElseIf f < 100 Then
        fstr = "00" & Trim$(Str(f))
    ElseIf f < 1000 Then
        fstr = "0" & Trim$(Str(f))
    Else
        fstr = Trim$(Str(f))
    End If
    FileName$ = "D:\MEDLINE Processed\medline17n" & fstr & ".xml"           'MEDLINE XML is pre-processed to reduce variation (e.g., "p less than" converted to "p<")
    lblProgress.Caption = "Processing File# " & fstr: lblProgress.Refresh
    fsize = FileLen(FileName$)
    FileNum1 = FreeFile
    Open FileName$ For Input As FileNum1
    bytesread = 0: recs = 0: abstr = 0
   
    While Not EOF(FileNum1)
        'Grab a chunk of text for processing
        If bytesread + BlockSize > fsize Then
            a$ = Input(fsize - bytesread, FileNum1)
            bytesread = bytesread + Len(a$)
        Else
            a$ = Input(BlockSize, FileNum1)
            bytesread = bytesread + Len(a$)
            b$ = vbNullString: endflag = 0
            While endflag = 0           'Extend block read to end of article (to avoid cutting key fields in half)
                b$ = b$ & Input(1, FileNum1)
                bytesread = bytesread + 1
                If bytesread = fsize Then endflag = 1
                If Right$(b$, 13) = "</RecordText>" Then endflag = 1
            Wend
            a$ = a$ & b$
        End If
        
        'PROCESS TEXT BLOCK
        endflag = 0: lastPDpos = 1
        RecSP$ = Split(a$, "</RecordText>")
        RecUB = UBound(RecSP$) - 1
        For RecCt = 0 To RecUB
            PubDatePos = InStr(RecSP$(RecCt), "</PubDate>")
            dat$ = Mid$(RecSP$(RecCt), PubDatePos - 8, 8)
            PMIDst = InStr(RecSP$(RecCt), "<PMID>")
            PMIDend = InStr(RecSP$(RecCt), "</PMID>")
            PMID = Val(Mid$(RecSP$(RecCt), PMIDst + 6, PMIDend - PMIDst - 6))
            JnameSt = InStr(RecSP$(RecCt), "<Journal>")
            JnameEnd = InStr(RecSP$(RecCt), "</Journal>")
            JName$ = Mid$(RecSP$(RecCt), JnameSt + 9, JnameEnd - JnameSt - 9)
            AbStart = InStr(RecSP$(RecCt), "<RecordText>")
            b$ = " " & Right$(RecSP$(RecCt), Len(RecSP$(RecCt)) - AbStart - 12) & " "
            If InStr(b$, "{") > 0 Then b$ = Replace(b$, "{", "(")       'standardize parentheticals
            If InStr(b$, "}") > 0 Then b$ = Replace(b$, "}", "(")
            If InStr(b$, "[") > 0 Then b$ = Replace(b$, "[", "(")
            If InStr(b$, "]") > 0 Then b$ = Replace(b$, "]", ")")
            If InStr(b$, "( ") > 0 Then b$ = Replace(b$, "( ", "(")
            If InStr(b$, "(") > 0 Then b$ = Replace(b$, "(", " (")      'add space on left (so that ratio statements are on the leftmost side of each word)
            If InStr(b$, "  (") > 0 Then b$ = Replace(b$, "  (", " (")  'strip double spaces
            If InStr(b$, " =") > 0 Then b$ = Replace(b$, " =", "=")
            If InStr(b$, "= ") > 0 Then b$ = Replace(b$, "= ", "=")
            If InStr(b$, "relative risk") > 0 Then b$ = Replace(b$, "relative risk", "RR")
            If InStr(b$, "risk ratio") > 0 Then b$ = Replace(b$, "risk ratio", "RR")
            If InStr(b$, "odds ratio") > 0 Then b$ = Replace(b$, "odds ratio", "OR")
            If InStr(b$, "hazard ratio") > 0 Then b$ = Replace(b$, "hazard ratio", "HR")
            If InStr(b$, "OR (OR)") > 0 Then b$ = Replace(b$, "OR (OR)", "OR")
            If InStr(b$, "RR (RR)") > 0 Then b$ = Replace(b$, "RR (RR)", "RR")
            If InStr(b$, "HR (HR)") > 0 Then b$ = Replace(b$, "HR (HR)", "HR")
            UnmodAbstr = b$
            LastPeriod = 0
            
            If AbStart > 0 Then
                sp$ = Split(b$, " ")
                UnmodSP$ = Split(UnmodAbstr$, " ")
                UB = UBound(sp$)
                For t = 0 To UB - 1
                    If Right$(sp$(t), 1) = "." Then LastPeriod = t
                    'look for (OR=, (RR=, (HR= pattern - these are high-confidence instances
                    ParenFlag = 0
                    ORflag = 0: RRflag = 0: HRflag = 0
                    LOR$ = Left$(sp$(t), 3)
                    d$ = Mid$(sp$(t), 4, 1)
                    If d$ Like "[;:,= ]" Then
                        If LOR$ = "(OR" Then
                            ParenFlag = 1: ORflag = 1: sp$(t) = Replace(sp$(t), "(OR", "(\R\")  'replace with standard pattern and set ORflag
                        End If
                        If LOR$ = "(RR" Then
                            ParenFlag = 1: RRflag = 1: sp$(t) = Replace(sp$(t), "(RR", "(\R\")  'replace with standard pattern and set RRflag
                        End If
                        If LOR$ = "(HR" Then
                            ParenFlag = 1: HRflag = 1: sp$(t) = Replace(sp$(t), "(HR", "(\R\")  'replace with standard pattern and set HRflag
                        End If
                    End If
                    If d$ Like "[;:, ]" Then
                        repl$ = "(\R\" & d$
                        sp$(t) = Replace(sp$(t), repl$, "(\R\=")    'standardize
                    End If
                    If ParenFlag = 1 Then       'A potential reporting item is on the left paren, so find the matching right paren
                        ORstring = vbNullString: ExamineFlag = 1   '(first word should have left paren)
                        u = t
                        While u < UB - 1
                            ORstring = ORstring & sp$(u) & " "      'extend to next paren or bracket to get full statement
                            If InStr(sp$(u), "(") > 0 Then          'nested parentheses. For each "(", demand one more ")"
                                ExamineFlag = ExamineFlag - 1
                            End If
                            If InStr(sp$(u), ")") > 0 Then
                                ExamineFlag = ExamineFlag + 1
                            End If
                            If ExamineFlag = 1 Then                 'found the matching right parenthetical, so drop out of loop
                                u = UB
                            End If
                            u = u + 1
                        Wend
                        'contains CI (confidence interval) or CL (confidence limit) acronym
                        lcor = LCase(ORstring)
                        If InStr(lcor$, "ci") = 0 And InStr(lcor$, "cl") = 0 And InStr(lcor$, "confidence interval") = 0 And InStr(lcor$, "confidence limit") = 0 Then ExamineFlag = 0
                    End If
                        
                    If ExamineFlag = 1 Then         'found a pattern, start processing
                        Context$ = vbNullString
                        If LastPeriod > 0 Then
                            u = LastPeriod + 1
                        Else
                            u = 0
                        End If
                        loopflag = 0
                        While loopflag = 0          'get surrounding context
                            Context = Context & UnmodSP$(u) & " "
                            If u < UB Then
                                If Right$(UnmodSP$(u), 1) = "." Then loopflag = 1
                            Else
                                loopflag = 1
                            End If
                            u = u + 1
                        Wend
                        'parse the odds ratio expression. First, reduce variation.
                        ORstring = Trim(ORstring)
                        If Right$(ORstring, 1) = "." Then ORstring = Left$(ORstring, Len(ORstring) - 1)
                        If Right$(ORstring, 1) = "," Then ORstring = Left$(ORstring, Len(ORstring) - 1)
                        If Right$(ORstring, 1) = ";" Then ORstring = Left$(ORstring, Len(ORstring) - 1)
                        If InStr(ORstring, "= ") > 0 Then ORstring = Replace(ORstring, "= ", "=")
                        If InStr(ORstring, "=OR") > 0 Then ORstring = Replace(ORstring, "=OR", "=")
                        If InStr(ORstring, "=RR") > 0 Then ORstring = Replace(ORstring, "=RR", "=")
                        If InStr(ORstring, "=HR") > 0 Then ORstring = Replace(ORstring, "=HR", "=")
                        d$ = Mid$(ORstring, 6, 1)       'sometimes double punctuation occurs
                        If d$ Like "[,:;-]" Then
                            ORstring = Replace(ORstring, "(\R\=" & d$, "(\R\=")
                        End If
                        modORstring = LCase(ORstring)
                        If InStr(modORstring, "�") > 0 Then modORstring = Replace(modORstring, "�", ".")        'decimal
                        If InStr(modORstring, "∼") > 0 Then modORstring = Replace(modORstring, "∼", "-")      'dash
                        If InStr(modORstring, "‒") > 0 Then modORstring = Replace(modORstring, "‒", "-")      'dash
                        If InStr(modORstring, "–") > 0 Then modORstring = Replace(modORstring, "–", "-")      'dash
                        If InStr(modORstring, "~") > 0 Then modORstring = Replace(modORstring, "~", "-")          'dash (assumed)
                        If InStr(modORstring, "≤") > 0 Then modORstring = Replace(modORstring, "≤", "<=")     '<=
                        If InStr(modORstring, "&lt;") > 0 Then modORstring = Replace(modORstring, "&lt;", "<")     '<
                        If InStr(modORstring, " ") > 0 Then modORstring = Replace(modORstring, " ", vbNullString)
                        If InStr(modORstring, " ") > 0 Then modORstring = Replace(modORstring, " ", vbNullString)
                        If InStr(modORstring, " to ") > 0 Then modORstring = Replace(modORstring, " to ", "-")
                        If InStr(modORstring, " approximately ") > 0 Then modORstring = Replace(modORstring, " approximately ", "-")
                        If InStr(modORstring, " of ") > 0 Then modORstring = Replace(modORstring, " of ", "=")
                        If InStr(modORstring, " with ") > 0 Then modORstring = Replace(modORstring, " with ", "|")
                        If InStr(modORstring, " for each ") > 0 Then modORstring = Replace(modORstring, " for each ", " per each ")
                        If InStr(modORstring, "p for trend") > 0 Then modORstring = Replace(modORstring, "p for trend", "p")
                        If InStr(modORstring, " for ") > 0 Then modORstring = Replace(modORstring, " for ", ", for ")
                        If InStr(modORstring, " in ") > 0 Then modORstring = Replace(modORstring, " in ", ", in ")
                        If InStr(modORstring, "percent") > 0 Then modORstring = Replace(modORstring, "percent", "%")
                        If InStr(modORstring, "per cent") > 0 Then modORstring = Replace(modORstring, "per cent", "%")
                        If InStr(modORstring, "confidence interval") > 0 Then modORstring = Replace(modORstring, "confidence interval", "ci")
                        If InStr(modORstring, "confidence limit") > 0 Then modORstring = Replace(modORstring, "confidence limit", "ci")
                        If InStr(modORstring, "cl") > 0 Then modORstring = Replace(modORstring, "cl", "ci")
                        If InStr(modORstring, "cis") > 0 Then modORstring = Replace(modORstring, "cis", "ci")
                        If InStr(modORstring, "ci (ci)") > 0 Then modORstring = Replace(modORstring, "ci (ci)", "ci")
                        If InStr(modORstring, "%ci") > 0 Then modORstring = Replace(modORstring, "%ci", "% ci")
                        If InStr(modORstring, " %") > 0 Then modORstring = Replace(modORstring, " %", "%")
                        If InStr(modORstring, " and 95% ci") > 0 Then modORstring = Replace(modORstring, " and 95% ci", ",95% ci")
                        If InStr(modORstring, "ci 95%") > 0 Then modORstring = Replace(modORstring, "ci 95%", "95% ci")
                        If InStr(modORstring, "ci 95") > 0 Then modORstring = Replace(modORstring, "ci 95", "95% ci")
                        If InStr(modORstring, "ci95%") > 0 Then modORstring = Replace(modORstring, "ci95%", "95% ci")
                        If InStr(modORstring, "ci95") > 0 Then modORstring = Replace(modORstring, "ci95", "95% ci")
                        If InStr(modORstring, "95%, ci") > 0 Then modORstring = Replace(modORstring, "95%, ci", "95% ci")
                        If InStr(modORstring, "95% ci - ") > 0 Then modORstring = Replace(modORstring, "95% ci - ", "95% ci=")
                        If InStr(modORstring, "95 ci") > 0 Then modORstring = Replace(modORstring, "95 ci", "95% ci")
                        If InStr(modORstring, "95%-ci") > 0 Then modORstring = Replace(modORstring, "95%-ci", "95% ci")
                        If InStr(modORstring, "95ci") > 0 Then modORstring = Replace(modORstring, "95ci", "95% ci")
                        If InStr(modORstring, "95%ci") > 0 Then modORstring = Replace(modORstring, "95%ci", "95% ci")
                        If InStr(modORstring, "95%; ci") > 0 Then modORstring = Replace(modORstring, "95%; ci", "95% ci")
                        If InStr(modORstring, "95% ci (95% ci)") > 0 Then modORstring = Replace(modORstring, "95% ci (95% ci)", "95% ci")
                        If InStr(modORstring, "ci (95)") > 0 Then modORstring = Replace(modORstring, "ci (95)", "95% ci")
                        If InStr(modORstring, "ci (95%)") > 0 Then modORstring = Replace(modORstring, "ci (95%)", "95% ci")
                        If InStr(modORstring, "ci, 95%") > 0 Then modORstring = Replace(modORstring, "ci, 95%", "95% ci")
                        If InStr(modORstring, "ci:") > 0 Then modORstring = Replace(modORstring, "ci:", "ci=")
                        If InStr(modORstring, "ci :") > 0 Then modORstring = Replace(modORstring, "ci :", "ci=")
                        If InStr(modORstring, "ci(") > 0 Then modORstring = Replace(modORstring, "ci(", "ci=")
                        If InStr(modORstring, "ci,") > 0 Then modORstring = Replace(modORstring, "ci,", "ci=")
                        If InStr(modORstring, "ci;") > 0 Then modORstring = Replace(modORstring, "ci;", "ci=")
                        If InStr(modORstring, "ci ") > 0 Then modORstring = Replace(modORstring, "ci ", "ci=")
                        If InStr(modORstring, " (95%") > 0 Then modORstring = Replace(modORstring, " (95%", ", 95%")
                        If InStr(modORstring, " (ci") > 0 Then modORstring = Replace(modORstring, " (ci", ", ci")
                        If InStr(modORstring, ") p=") > 0 Then modORstring = Replace(modORstring, ") p=", "), p=")
                        If InStr(modORstring, ")p=") > 0 Then modORstring = Replace(modORstring, ")p=", "), p=")
                        If InStr(modORstring, " and p=") > 0 Then modORstring = Replace(modORstring, " and p=", "p=")
                        If InStr(modORstring, "p-value") > 0 Then modORstring = Replace(modORstring, "p-value", "p")
                        If InStr(modORstring, "p value") > 0 Then modORstring = Replace(modORstring, "p value", "p")
                        If InStr(modORstring, "p(trend)") > 0 Then modORstring = Replace(modORstring, "p(trend)", "p")
                        If InStr(modORstring, "p for interaction") > 0 Then modORstring = Replace(modORstring, "p for interaction", "p")
                        If InStr(modORstring, "==") > 0 Then modORstring = Replace(modORstring, "==", "=")
                        If InStr(modORstring, ",,") > 0 Then modORstring = Replace(modORstring, ",,", ",")
                        If InStr(modORstring, ";,") > 0 Then modORstring = Replace(modORstring, ";,", ",")
                        
                        'if multiple OR statments, split the string so the next OR in the main text so it can be found afterwards
                        If InStr(modORstring, " and or=") > 0 Or InStr(modORstring, "; or=") > 0 Or InStr(modORstring, " and rr=") > 0 Or InStr(modORstring, "; rr=") > 0 Or InStr(modORstring, " and hr=") > 0 Or InStr(modORstring, "; hr=") > 0 Then
                            u = t
                            While u < UB - 1
                                If (sp$(u) = "and" Or Right$(sp$(u), 1) = ";") Then
                                    If Left(sp$(u + 1), 3) = "OR=" Or Left(sp$(u + 1), 3) = "RR=" Or Left(sp$(u + 1), 3) = "HR=" Then
                                        sp$(u + 1) = "(" & sp$(u + 1)   'add paren so next OR can be found
                                        u = UB
                                    End If
                                End If
                                u = u + 1
                            Wend                                                                          'since we've already established it's there,
                            If InStr(modORstring, " or=") > 0 Then sp2$ = Split(modORstring, " or=")      'take the next OR statement out of this one
                            If InStr(modORstring, " rr=") > 0 Then sp2$ = Split(modORstring, " rr=")      'take the next RR statement out of this one
                            If InStr(modORstring, " hr=") > 0 Then sp2$ = Split(modORstring, " hr=")      'take the next HR statement out of this one
                            modORstring = sp2$(0) & ")"
                        End If
                                                
                        u = Len(modORstring)
                        For v = 1 To u
                            d$ = Mid$(modORstring, v, 5)
                            If d$ Like "0,###" Then           'leftmost zero indicates it's probably a decimal
                                repl$ = Replace(d$, ",", ".")
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "#,###" Then           'get rid of commas in thousands
                                repl$ = Replace(d$, ",", vbNullString)
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "# ci=" Then           'no delimiter for generic CI
                                repl$ = Replace(d$, " ci=", ", ci=")
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "# 95%" Then           'no delimiter for 95% CI
                                repl$ = Replace(d$, " 95%", ", 95%")
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "#,##[-,; )]" Then           'change European commas to decimals, based on delimiters
                                repl$ = Left$(d$, 1) & "." & Right$(d$, 3)     'change numeric "," but not delimiter ","
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            d$ = Mid$(modORstring, v, 4)
                            If d$ Like "#,#[-,; )]" Then           'change European commas to decimals, based on delimiters
                                repl$ = Left$(d$, 1) & "." & Right$(d$, 2)     'change numeric "," but not delimiter ","
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "# p=" Then           'no delimiter
                                repl$ = Replace(d$, " p=", ", p=")
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                            If d$ Like "# p<" Then           'no delimiter
                                repl$ = Replace(d$, " p<", ", p<")
                                modORstring = Replace(modORstring, d$, repl$)
                            End If
                        Next v
                        
                        modORstring = Replace(modORstring, " ", vbNullString)
                        If InStr(modORstring, ",") > 0 Then modORstring = Replace(modORstring, ",", "|")
                        If InStr(modORstring, ":") > 0 Then modORstring = Replace(modORstring, ":", "|")
                        If InStr(modORstring, ";") > 0 Then modORstring = Replace(modORstring, ";", "|")
                        While InStr(modORstring, "||") > 0
                            modORstring = Replace(modORstring, "||", "|")
                        Wend
                        If InStr(modORstring, "(") > 0 Then modORstring = Replace(modORstring, "(", vbNullString)
                        If InStr(modORstring, ")") > 0 Then modORstring = Replace(modORstring, ")", vbNullString)
                        If InStr(modORstring, "95%ci=") > 0 Then modORstring = Replace(modORstring, "95%ci=", vbNullString)
                        If InStr(modORstring, "ci=") > 0 Then modORstring = Replace(modORstring, "ci=", vbNullString)
                        
                        OddsR$ = vbNullString: CIlower$ = vbNullString: CIupper$ = vbNullString
                        logoddsR = 0: logoddsL = 0: logoddsU = 0: Middle = 0: diff = 0
                        
                        sp2$ = Split(modORstring, "|")
                        UB2 = UBound(sp2$)
                        If UB2 > 0 Then
                            For v = 0 To UB2
                                If Left$(sp2$(v), 3) = "for" Then   'these frequently intervene between OR CI pairs
                                    modORstring = Replace(modORstring, ("|" & sp2$(v)), vbNullString)
                                    v = UB2 + 1
                                    sp2$ = Split(modORstring, "|")
                                    UB2 = UBound(sp2$)
                                End If
                            Next v
                            If InStr(sp2$(0), "per") > 0 Then           'frequently the OR will be expressed in terms of per X
                                NewWord$ = Left(sp2$(0), InStr(sp2$(0), "per") - 1)
                                modORstring = Replace(modORstring, sp2$(0), NewWord$)
                                sp2$(0) = NewWord$
                            End If
                            If Left(sp2$(0), 4) = "\r\=" Then
                                OddsR$ = Replace(sp2$(0), "\r\=", vbNullString)  'First array element is reportable item (renamed \r\)
                            End If
                        End If
                        If UB2 = 1 Then                        'second might be CI
                            If InStr(sp2$(1), "-") > 0 Then    'the lower and upper are divided by a "-" sign
                                sp3$ = Split(sp2$(1), "-")
                                CIlower$ = sp3$(0)
                                CIupper$ = sp3$(1)
                            End If
                        End If
                        If UB2 >= 2 Then     'the lower and upper are divided by some other sign (e.g., comma or sem)
                            If IsNumeric(sp2$(1)) = True And IsNumeric(sp2$(2)) = True Then        'if the 2 values are numbers, then likely the delimiter was not a "-"
                                CIlower$ = sp2$(1)
                                CIupper$ = sp2$(2)
                            Else
                                If InStr(sp2$(1), "-") > 0 Then    'if "-" is in the 2nd field, then the final value is likely a p-value
                                    sp3$ = Split(sp2$(1), "-")
                                    CIlower$ = sp3$(0)
                                    CIupper$ = sp3$(1)
                                End If
                            End If
                        End If
                        pval$ = vbNullString
                        For v = 1 To UB2
                        'NOTE: need to refine p-value extraction
                            If InStr(sp2$(v), "p<") > 0 Or InStr(sp2$(v), "p=") > 0 Or InStr(sp2$(v), "p>") > 0 Then
                                If pval$ = vbNullString Then pval$ = sp2$(v)        'Take the 1st p-value only in case there are multiple
                                'sometimes the p-value comes second, so if no CI limits have been found yet then check that possibility
                                If CIlower = "" And CIupper = "" And v = 1 And UB2 >= 2 Then
                                    If InStr(sp2$(2), "-") > 0 Then    'if "-" is in the next field, then split
                                        sp3$ = Split(sp2$(2), "-")
                                        CIlower$ = sp3$(0)
                                        CIupper$ = sp3$(1)
                                    Else
                                        If UB2 >= 3 Then                'else, if a 3rd field is present, then use lower=2nd, upper=3rd
                                            CIlower$ = sp2$(2)
                                            CIupper$ = sp2$(3)
                                        End If
                                    End If
                                End If
                            End If
                        Next v
                        If pval$ <> vbNullString Then           'strip extra chars off p-val (if any)
                            If Left$(pval$, 2) <> "p=" And Left$(pval$, 2) <> "p<" And Left$(pval$, 2) <> "p>" Then
                                If InStr(pval$, "p=") > 0 Then pval$ = Right$(pval$, Len(pval$) - InStr(pval$, "p=") + 1)
                                If InStr(pval$, "p>") > 0 Then pval$ = Right$(pval$, Len(pval$) - InStr(pval$, "p>") + 1)
                                If InStr(pval$, "p<") > 0 Then pval$ = Right$(pval$, Len(pval$) - InStr(pval$, "p<") + 1)
                            End If
                            pval$ = Replace(pval$, " ", vbNullString)
                            If InStr(pval$, "p=10(-") > 0 Then pval$ = Replace(pval$, "p=10(-", "p=1x10(-")
                            pval$ = Replace(pval$, "(", vbNullString)
                            pval$ = Replace(pval$, ")", vbNullString)
                            pval$ = Replace(pval$, "/", vbNullString)
                            pval$ = Replace(pval$, "`", vbNullString)
                            pval$ = Replace(pval$, "*", vbNullString)
                            If InStr(pval$, "*10") > 0 Then pval$ = Replace(pval$, "*10", "x10")
                            If InStr(pval$, "x10") > 0 Then pval$ = Replace(pval$, "x10", "E") 'convert exponent statements (e.g., 1.2x10(-4) to 1.2E-4)
                            If InStr(pval$, "e-") > 0 Then pval$ = Replace(pval$, "e-", "E-")  'uppercase will save it from the trimming routine below
                            If InStr(pval$, ".") = 0 And InStr(pval$, ",") > 0 Then pval$ = Replace(pval$, ",", ".")   'if there is no decimal but a comma, then it may be European convention
                            
                            If InStr(pval$, "p=.") > 0 Then pval$ = Replace(pval$, "p=.", "p=0.")
                            If InStr(pval$, "p<.") > 0 Then pval$ = Replace(pval$, "p<.", "p<0.")
                            If InStr(pval$, "p<=.") > 0 Then pval$ = Replace(pval$, "p<=.", "p<=0.")
                            If InStr(pval$, "p>.") > 0 Then pval$ = Replace(pval$, "p>.", "p>0.")
                            lenp = Len(pval$)
                            For v = 4 To lenp
                                If Mid$(pval$, v, 1) Like "[a-z]" Then
                                    pval$ = Left(pval$, v - 1)
                                    v = lenp
                                End If
                            Next v
                            If Right$(pval$, 1) = "." Then pval$ = Left$(pval$, Len(pval$) - 1)
                        End If
                        
                        If IsNumeric(OddsR) = False Then OddsR = vbNullString    'check to make sure the value has been isolated without extraneous characters
                        If IsNumeric(CIlower) = False Then CIlower = vbNullString
                        If IsNumeric(CIupper) = False Then
                            If InStr(CIupper, "%") = 0 And InStr(CIupper, "&") = 0 Then         '% and & give VAL error
                                If Val(CIupper) > 0 Then
                                    CIupper = Str$(Val(CIupper))   'sometimes trailing words are included after the upper limit - if so, strip them out by taking VAL
                                Else
                                    CIupper = vbNullString
                                End If
                            Else
                                CIupper = vbNullString
                            End If
                        End If
                        ErrFlag = vbNullString
                        If Val(OddsR) <= 0 Then OddsR = vbNullString
                        If Val(CIlower) < 0 Then CIlower = vbNullString     'Lower CI might be zero
                        If Val(CIupper) <= 0 Then CIupper = vbNullString
                        If Val(CIlower) < 0 Or Val(CIupper) < 0 Then ErrFlag = ErrFlag & "Neg CI "
                        ProcessedFlag = 0
                        If OddsR$ <> vbNullString And CIlower$ <> vbNullString And CIupper$ <> vbNullString Then
                            'flag potential errors
                            If Val(CIupper) < Val(CIlower) Then ErrFlag = ErrFlag & "L>U "
                            If Val(CIlower) = 0 Then ErrFlag = ErrFlag & "CI(L)=0 "
                            
                            logoddsR = Log(Val(OddsR)) / Log(10)
                            If Val(CIlower) > 0 Then
                                logoddsL = Log(Val(CIlower)) / Log(10)
                            Else
                                logoddsL = 0
                            End If
                            logoddsU = Log(CIupper) / Log(10)
                            Middle = (logoddsU + logoddsL) / 2
                            'round based on sig figs and check and see if it's better to round up or down
                            ReportedNum = OddsR$
                            If Middle < 15 Then             'huge numbers will cause overflow
                                CalcNum = 10 ^ Middle
                                If CalcNum > 0 Then Call Choose_round_up_or_down(ReportedNum$, CalcNum)
                                OddsRdiff = Val(OddsR$)
                                diff = OddsRdiff - CalcNum
                                Call Calculate_maximum_interval(CIlower, CIupper, OddsR, LowestOR, HighestOR, MinDiff)
                                If diff = 0 And MinDiff > 0 Then
                                    MinDiff = 0
                                End If
                                ProcessedFlag = 1
                            Else
                                ErrFlag = ErrFlag & "too big "
                            End If
                        End If
                        If ORflag = 1 Then                          'Restore former statement for reporting purposes
                            ORstring = Replace(ORstring, "(\R\=", "(OR=")
                            modORstring = Replace(modORstring, "\r\=", "or=")
                        End If
                        If RRflag = 1 Then
                            ORstring = Replace(ORstring, "(\R\=", "(RR=")
                            modORstring = Replace(modORstring, "\r\=", "rr=")
                        End If
                        If HRflag = 1 Then
                            ORstring = Replace(ORstring, "(\R\=", "(HR=")
                            modORstring = Replace(modORstring, "\r\=", "hr=")
                        End If
                        If ProcessedFlag = 1 Then       'Was able to extract all parameters
                            Print #FileNum2, PMID & vbTab & ORstring & vbTab & modORstring & vbTab & pval$ & vbTab & CIlower & vbTab & OddsR & vbTab & CIupper & vbTab & 10 ^ Middle & vbTab & LowestOR$ & vbTab & CalcNum & vbTab & HighestOR$ & vbTab & MinDiff & vbTab & ErrFlag & vbTab & JName$ & vbTab & Context$
                        Else                            'did not get all parameters - output it to figure out why
                            Print #FileNum2, PMID & vbTab & ORstring & vbTab & modORstring & vbTab & pval$ & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ErrFlag & vbTab & JName$ & vbTab & Context$
                        End If
                        ExamineFlag = 0
                    End If
                Next t
            End If      'if abstart > 0
        Next RecCt
        inp$ = " ": b$ = vbNullString
    Wend
lblProgress.Caption = vbNullString
Close FileNum1
Next f
Close FileNum2
frmProcessing.MousePointer = vbDefault

End Sub
Public Sub Calculate_maximum_interval(CIlower$, CIupper$, OddsR$, LowestOR$, HighestOR$, MinDiff@)
'Assuming that the actual interval range might be slightly different due to rounding (e.g., was really 1.15 but was reported as 1.2),
'shift both CI to the lowest possible and the highest possible, then recalulate ORs
Dim sp$(), RoundedUpL@, RoundedDownL@, RoundedUpU@, RoundedDownU@, LOR@, UOR@, ReportedOR@
Dim t&, MaxND&, ND1&, ND2&

'The # of decimals SHOULD be the same for both intervals, but just in case
If InStr(1, CIlower$, ".") > 0 Then
    sp$ = Split(CIlower$, ".")
    ND1 = Len(sp$(1))
Else
    ND1 = 0
End If
If InStr(1, CIupper$, ".") > 0 Then
    sp$ = Split(CIupper$, ".")
    ND2 = Len(sp$(1))
Else
    ND2 = 0
End If
If ND1 >= ND2 Then
    MaxND = ND1
Else
    MaxND = ND2
End If
'calculate lowest possible OR (due to rounding)
RoundedDownL = Val(CIlower$) - 0.5 * 10 ^ (ND1 * -1)
RoundedDownU = Val(CIupper$) - 0.5 * 10 ^ (ND2 * -1)
If RoundedDownL > 0 Then
    LOR = ((Log(RoundedDownL) / Log(10)) + (Log(RoundedDownU) / Log(10))) / 2
    LOR = 10 ^ LOR
    LOR = Int(LOR * (10 ^ MaxND)) / (10 ^ MaxND)            'round down to give max benefit of doubt
    LowestOR$ = Str$(LOR)
Else
    LowestOR$ = "0"
End If

'calculate highest possible OR (due to rounding)
RoundedUpL = Val(CIlower$) + 0.5 * 10 ^ (ND1 * -1)
RoundedUpU = Val(CIupper$) + 0.5 * 10 ^ (ND2 * -1)
If RoundedUpL > 0 Then
    UOR = ((Log(RoundedUpL) / Log(10)) + (Log(RoundedUpU) / Log(10))) / 2
    UOR = 10 ^ UOR
    UOR = Int(UOR * (10 ^ MaxND) + 0.5) / (10 ^ MaxND)      'round up to give max benefit of doubt
    HighestOR$ = Str$(UOR)
Else
    HighestOR = "0"
End If

ReportedOR@ = Val(OddsR$)
If ReportedOR >= LOR And ReportedOR <= UOR Then
    MinDiff = 0
Else
    If Abs(ReportedOR - LOR) <= Abs(ReportedOR - UOR) Then
        MinDiff = Abs(ReportedOR - LOR)
    Else
        MinDiff = Abs(ReportedOR - UOR)
    End If
End If
End Sub
Public Sub Choose_round_up_or_down(ReportedNum$, CalcNum@)
'Based on the # of decimals and sig figs to the right of the decimal in the reported OR, try rounding up and down
'and go with whichever is closer
Dim sp$(), RoundedUp@, RoundedDown@
Dim t&, sigfigs%, decimalpoints%, SF1&, SF2&

Call Calc_Sig_Figs(ReportedNum$, sigfigs)       'get sig figs for the reported OR/RR
RoundedUp = FormatSF(CalcNum, sigfigs)          'use this # of sig figs for the calculated #
RoundedDown = FormatSFdown(CalcNum, sigfigs)
If Abs(RoundedUp - Val(ReportedNum$)) > Abs(RoundedDown - Val(ReportedNum$)) Then
    CalcNum = RoundedDown
Else
    CalcNum = RoundedUp
End If
End Sub
Public Sub Calc_Sig_Figs(ReportedNum$, sigfigs%)
Dim sp$(), t&
If InStr(1, ReportedNum$, ".") > 0 Then
    sp$ = Split(ReportedNum$, ".")
    If sp$(0) = "0" Or sp$(0) = " " Then    'if less than one, find out where the #s begin (e.g., 0.0054 has two sig figs, 0.00540 has three)
        For t = 1 To Len(sp$(1))
            If Mid$(sp$(1), t, 1) <> "0" Then
                sigfigs = Len(sp$(1)) - (t - 1)
                t = Len(sp$(1)) + 1
            End If
        Next t
    Else                    'is > 1 so add both sides
        sigfigs = Len(sp$(0)) + Len(sp$(1))
    End If
Else
    sigfigs = Len(ReportedNum$)
End If
End Sub
'Returns input number rounded to specified number of significant figures.
Function FormatSF(dblInput@, intSF As Integer) As String
Dim intCorrPower As Integer         'Exponent used in rounding calculation
Dim intSign As Integer              'Holds sign of dblInput since logs are used in calculations
 
'-- Store sign of dblInput --
intSign = Sgn(dblInput)
'-- Calculate exponent of dblInput --
intCorrPower = Int(Log10(Abs(dblInput)))
FormatSF = Round(dblInput * 10 ^ ((intSF - 1) - intCorrPower))   'integer value with no sig fig
FormatSF = FormatSF * 10 ^ (intCorrPower - (intSF - 1))         'raise to original power
'-- Reconsitute final answer --
FormatSF = FormatSF * intSign
'-- Answer sometimes needs padding with 0s --
If InStr(FormatSF, ".") = 0 Then
    If Len(FormatSF) < intSF Then
        FormatSF = Format(FormatSF, "##0." & String(intSF - Len(FormatSF), "0"))
    End If
End If
If intSF > 1 And Abs(FormatSF) < 1 Then
    Do Until Left(Right(FormatSF, intSF), 1) <> "0" And Left(Right(FormatSF, intSF), 1) <> "."
        FormatSF = FormatSF & "0"
    Loop
End If
End Function
'Returns input number rounded DOWN to specified number of significant figures.
Function FormatSFdown(dblInput@, intSF As Integer) As String
Dim intCorrPower As Integer         'Exponent used in rounding calculation
Dim intSign As Integer              'Holds sign of dblInput since logs are used in calculations
 
'-- Store sign of dblInput --
intSign = Sgn(dblInput)
'-- Calculate exponent of dblInput --
intCorrPower = Int(Log10(Abs(dblInput)))
FormatSFdown = Int(dblInput * 10 ^ ((intSF - 1) - intCorrPower))    'integer value with no sig fig - round DOWN
FormatSFdown = FormatSFdown * 10 ^ (intCorrPower - (intSF - 1))         'raise to original power
'-- Reconsitute final answer --
FormatSFdown = FormatSFdown * intSign
'-- Answer sometimes needs padding with 0s --
If InStr(FormatSFdown, ".") = 0 Then
    If Len(FormatSFdown) < intSF Then
        FormatSFdown = Format(FormatSFdown, "##0." & String(intSF - Len(FormatSFdown), "0"))
    End If
End If
If intSF > 1 And Abs(FormatSFdown) < 1 Then
    Do Until Left(Right(FormatSFdown, intSF), 1) <> "0" And Left(Right(FormatSFdown, intSF), 1) <> "."
        FormatSFdown = FormatSFdown & "0"
    Loop
End If
End Function
'Calculate Log to the Base 10
Function Log10(x)
   Log10 = Log(x) / Log(10#)
End Function
Public Sub Convert_Month(mon$)
If Len(mon$) = 7 Then           'Range (e.g. 1992 Jan-Feb)
    If Mid$(mon$, 4, 1) = "-" Then mon$ = Left$(mon$, 3)
End If
Select Case mon$
Case "Jan": mon$ = "01"
Case "Feb": mon$ = "02"
Case "Mar": mon$ = "03"
Case "Apr": mon$ = "04"
Case "May": mon$ = "05"
Case "Jun": mon$ = "06"
Case "Jul": mon$ = "07"
Case "Aug": mon$ = "08"
Case "Sep": mon$ = "09"
Case "Oct": mon$ = "10"
Case "Nov": mon$ = "11"
Case "Dec": mon$ = "12"
Case Else
    mon$ = "unknown"
End Select
End Sub

