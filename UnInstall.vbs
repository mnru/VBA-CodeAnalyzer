'=========================================================================================
'Tear down ExcelAddIn
'=========================================================================================

Const FILLE_NAME="VBACodeAnalyzer.xlam"

Call Exec

Sub Exec()
    Dim objExcel
    Dim strAdPath
    Dim strMyPath
    Dim strAdCp
    Dim strMyCp
    Dim objFileSys
    Dim oAdd
    
	on error resume next

    '-- CreateObject
    Set objExcel = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    '-- Set Path
    strAdPath = objExcel.Application.UserLibraryPath
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strAdCp = objFileSys.BuildPath(strAdPath, FILLE_NAME)
    strMyCp = objFileSys.BuildPath(strMyPath, FILLE_NAME)
    '-- Delete from Excel 
    objExcel.Workbooks.Add
    Set oAdd = objExcel.AddIns.Add(strAdCp,True)
    oAdd.Installed = False
    objExcel.Quit
   '-- Delete File
    objFileSys.DeleteFile strAdCp

    '-- Free Object
    Set objExcel = Nothing
    Set objFileSys = Nothing
    
    MsgBox "Complete!"
	
End Sub