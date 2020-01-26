Attribute VB_Name = "modAnalyze"
Option Base 0
Sub analyzeOtherBook(thisRibbon)
    
    pn = Application.GetOpenFilename("excel macro book,*.xlsm,all file,*.*", , "select workbook to analyze")
    If pn = "False" Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    bn = fso.getFilename(pn)
    Workbooks.Open (pn)
    Call analyzeCode(thisRibon, True)
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks(bn).Close
    Application.DisplayAlerts = True
End Sub

Sub analyzeCode(thisRibbon, Optional otherbook = False)
    bn = ActiveWorkbook.Name
    
    Dim currentRow        As Long
    Dim currentSummaryRow As Long
    Dim procCnt           As Long
    Dim propertyCnt       As Long
    Dim mdlName           As String
    Dim procName
    Dim procLineNum       As Long
    Dim mdlLineNum        As Long
    Dim declareLineNum    As Long
    Dim defInfo
    Dim lineCnt           As Long
    Workbooks(bn).Activate
    Sheets.Add
    sn = ActiveSheet.Name
    Range("a1") = bn
    currentRow = 3
    currentSummaryRow = 3
    arynum = 11
    aryTittle = Array("module", "(type)", "lines", "proc", "name", "arg", "return type", "", "def line", "lines", "comment", "colon", "signature")
    arySummaryTitle = Array("module", "type", "fun/sub", "(property)", "total lines", "(declaration)", "(procedures)")
    arynum = lenAry(aryTittle) - LBound(aryTittle) + 1
    Worksheets(sn).Cells(currentRow, 9).Resize(1, arynum) = aryTittle
    Worksheets(sn).Cells(currentRow, 1).Resize(1, 7) = arySummaryTitle
    currentRow = currentRow + 1
    currentSummaryRow = currentSummaryRow + 1
    Workbooks(bn).Activate
    
    On Error Resume Next
    For Each cmp In ActiveWorkbook.VBProject.VBComponents
        With cmp.CodeModule
            If .CountOfLines > 0 Then
                Set dic = CreateObject("Scripting.Dictionary")
                mdlName = cmp.Name
                mdlType = getModType(cmp)
                mdlLineNum = .CountOfLines
                declareLineNum = .CountOfDeclarationLines
                procCnt = 0
                codelineCnt = 0
                procName = ""
                For lineCnt = 1 To .CountOfLines
                    If procName <> .ProcOfLine(lineCnt, 0) Then
                        procName = .ProcOfLine(lineCnt, 0)
                        procLineNum = tryToGetProcLineNum(cmp, procName, 0)
                        If procLineNum = 0 Then
                            Call dic.Add(procName, Null)
                        Else
                            defInfo = getDef(cmp, procName)
                            Call writeData(procCnt, mdlName, mdlType, procLineNum, procName, defInfo, sn, currentRow, codelineCnt)
                        End If
                    End If
                Next lineCnt
                propertyCnt = 0
                For Each procName In dic.keys
                    For knd = 1 To 3
                        procLineNum = tryToGetProcLineNum(cmp, procName, knd)
                        If procLineNum <> 0 Then
                            defInfo = getDef(cmp, procName, knd)
                            Call writeData(procCnt, mdlName, mdlType, procLineNum, procName, defInfo, sn, currentRow, codelineCnt)
                            propertyCnt = propertyCnt + 1
                        End If
                    Next knd
                Next
                arySummary = Array(mdlName, mdlType, procCnt, propertyCnt, mdlLineNum, declareLineNum, codelineCnt)
                
                Worksheets(sn).Cells(currentSummaryRow, 1).Resize(1, 7) = arySummary
                currentSummaryRow = currentSummaryRow + 1
            End If
        End With
    Next cmp
    
    Call prettyDisplay(sn)
    Range("A1").WrapText = False
    
    
    If otherbook Then
        Application.DisplayAlerts = False
        Worksheets(sn).Copy
        Application.DisplayAlerts = True
    End If
End Sub

Sub prettyDisplay(sn)
    With Sheets(sn).Cells
        .WrapText = False
        .Columns.AutoFit
        .Rows.AutoFit
        .WrapText = True
        .Columns.AutoFit
        .Rows.AutoFit
    End With
    
    With Sheets(sn)
        Call .ListObjects.Add(xlSrcRange, .Cells(3, 1).CurrentRegion, , xlYes)
        Call .ListObjects.Add(xlSrcRange, .Cells(3, 9).CurrentRegion, , xlYes)
        
    End With
End Sub

Function analyzeSignature(str, fn)
    Dim reg, mc, n0, n1, n2, p1, p2, p3, str0, ary
    n0 = Len(str)
    n1 = InStr(str, fn & "(")
    n2 = n1 + Len(fn)
    p1 = Trim(Left(str, n1 - 1))
    'p1 = Left(str, n2)
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "\)\s+As\s+([^\(\)\s]+(\(\))?)s*$"
        
        .IgnoreCase = False
        .Global = False
    End With
    
    Set mc = reg.Execute(str)
    If mc.Count = 0 Then
        n3 = 1
        p3 = ""
       ' p3 = Right(str, 1)
    Else
        n3 = Len(mc(0))
        p3 = mc(0).submatches(0)
        'p3 = mc(0)
    End If
    
    str0 = Mid(str, n2 + 1, n0 - n2 - n3)
    
    'p2 = splitQuotation(str0, ",", vbLf)
    p2s = rejoinWithQuotation(str0, ",", vbLf)
    p2 = p2s(0)
    
    ret = Array(p1, p2, p3)
    
    analyzeSignature = ret
    Set reg = Nothing
    Set mc = Nothing
End Function

Function rejoinWithQuotation(str1, dlm0, Optional dlm1 = "", Optional break As Boolean = True)
    If dlm1 <> "" Then break = False
    
    Dim ret  As String
    Dim tmp  As String
    Dim flag As Boolean
    Dim num, cnt, qcnt
    xs = Split(str1, dlm0)
    num = lenAry(xs)
    ret = ""
    tmp = ""
    cnt = 0
    qcnt = 0
    
    For i = LBound(xs) To UBound(xs)
        cnt = cnt + 1
        If tmp <> "" Then tmp = tmp & dlm0
        tmp = tmp & xs(i)
        qcnt = qcnt + countStr(xs(i), """")
        If qcnt Mod 2 = 0 Then
            tmp = Trim(tmp)
            If ret <> "" Then ret = ret & dlm1
            ret = ret & tmp
            tmp = ""
            If break Then Exit For
        End If
    Next i
    flag = IIf(cnt = num, False, True)
    rejoinWithQuotation = Array(ret, flag)
    
End Function


Function splitQuotation(str1, dlm0, dlm1)
    
    ret = ""
    tmp = ""
    xs = Split(str1, dlm0)
    For Each x In xs
        If tmp <> "" Then
            tmp = tmp & dlm0 & x
            n0 = countStr(tmp, """")
            If n0 Mod 2 = 0 Then
                If ret <> "" Then ret = ret & dlm1
                ret = ret & Trim(tmp)
                tmp = ""
            End If
        Else
            n = countStr(x, """")
            If n Mod 2 = 0 Then
                If ret <> "" Then ret = ret & dlm1
                ret = ret & Trim(x)
            Else
                tmp = x
            End If
        End If
    Next x
    splitQuotation = ret
    
End Function
Function breakQuotation(str1, dlm0)
    Dim ret  As String
    Dim flag As Boolean
    Dim tmp  As String
    ret = ""
    tmp = ""
    Dim num, cnt
    cnt = 0
    xs = Split(str1, dlm0)
    num = lenAry(xs)
    For Each x In xs
        cnt = cnt + 1
        If tmp <> "" Then
            tmp = tmp & dlm0 & x
            n0 = countStr(tmp, """")
            If n0 Mod 2 = 0 Then
                If ret <> "" Then ret = ret & dlm0
                ret = ret & Trim(tmp)
                tmp = ""
                Exit For
            End If
        Else
            n = countStr(x, """")
            If n Mod 2 = 0 Then
                If ret <> "" Then ret = ret & dlm0
                ret = ret & Trim(x)
                Exit For
            Else
                tmp = x
            End If
        End If
    Next x
    flag = IIf(cnt = num, False, True)
    breakQuotation = Array(ret, flag)
    
End Function

Function countStr(str1, dlm0)
    ret = Len(str1) - Len(Replace(str1, dlm0, ""))
    countStr = ret
End Function

Function tryToGetProcLineNum(cmp, procName, Optional knd = 0)
    On Error Resume Next
    ret = 0
    ret = cmp.CodeModule.ProcCountLines(procName, knd)
    tryToGetProcLineNum = ret
End Function

Function getModType(cmp)
    Dim ret
    Select Case cmp.Type
        Case 1: ret = "Std"
        Case 2: ret = "Cls"
        Case 3: ret = "Frm"
        Case Else: ret = ""
    End Select
    getModType = ret
End Function

Function getKndName(num)
    Dim ret
    Select Case num
        Case 1: ret = "Let"
        Case 2: ret = "Set"
        Case 3: ret = "Get"
        Case Else: ret = ""
    End Select
    getKndName = ret
End Function

Function getDef(cmp, procName, Optional knd = 0)
    With cmp.CodeModule
        lineDef = .procBodyLine(procName, knd)
        lineEnd = .ProcCountLines(procName, knd) + .ProcStartLine(procName, knd)
        cnt = 0
        Do While lineDef + cnt < lineEnd
            strLine = Trim(.Lines(lineDef + cnt, 1))
            cnt = cnt + 1
            If strLine Like "* _" Then
                ret = ret & Left(strLine, Len(strLine) - 1)
            Else
                ret = ret & strLine
                Exit Do
            End If
        Loop
        
        Dim x0, x1
        x0 = breakQuotation(ret, "'")
        x1 = breakQuotation(x0(0), ":")
'        x0 = rejoinWithQuotation(ret, "'")
'        x1 = rejoinWithQuotation(x0(0), ":")
'
        
        flgComment = x0(1)
        flgColon = x1(1)
        ret = x1(0)
    End With
    ret = Trim(ret)
    getDef = Array(lineDef, cnt, flgComment, flgColon, ret)
End Function

Sub writeData(procCnt, mdlName, mdlType, procLineNum, procName, ByVal defInfo, sn, currentRow, codelineCnt)
    Dim parts, df
    procCnt = procCnt + 1
    df = defInfo
    parts = analyzeSignature(df(4), procName)
    ary = Array(mdlName, mdlType, procLineNum, parts(0), procName, parts(1), parts(2), "", df(0), df(1), df(2), df(3), df(4))
    arynum = lenAry(ary)
    Worksheets(sn).Cells(currentRow, 9).Resize(1, arynum) = ary
    codelineCnt = codelineCnt + procLineNum
    currentRow = currentRow + 1
End Sub

Private Function lenAry(ary)
    lenAry = UBound(ary) - LBound(ary) + 1
End Function



