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
    aryTittle = Array("No", "module", "lines", "fun/sub", "def line", "lines", "signature")
    arySummaryTitle = Array("module", "type", "fun/sub", "(property)", "total lines", "(declaration)", "(procedures)")
    Worksheets(sn).Cells(currentRow, 9).Resize(1, 7) = aryTittle
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
                            Call writeData(procCnt, mdlName, procLineNum, procName, defInfo, sn, currentRow, codelineCnt)
                        End If
                    End If
                Next lineCnt
                propertyCnt = 0
                For Each procName In dic.keys
                    For knd = 1 To 3
                        procLineNum = tryToGetProcLineNum(cmp, procName, knd)
                        If procLineNum <> 0 Then
                            defInfo = getDef(cmp, procName, knd)
                            Call writeData(procCnt, mdlName, procLineNum, getKndName(knd) & " " & procName, defInfo, sn, currentRow, codelineCnt)
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
    If otherbook Then
        Application.DisplayAlerts = False
        Worksheets(sn).Copy
        Application.DisplayAlerts = True
    End If
End Sub

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
    End With
    getDef = Array(lineDef, cnt, ret)
End Function

Sub writeData(procCnt, mdlName, procLineNum, procName, defInfo, sn, currentRow, codelineCnt)
    procCnt = procCnt + 1
    ary = Array(procCnt, mdlName, procLineNum, procName, defInfo(0), defInfo(1), defInfo(2))
    Worksheets(sn).Cells(currentRow, 9).Resize(1, 7) = ary
    codelineCnt = codelineCnt + procLineNum
    currentRow = currentRow + 1
End Sub
