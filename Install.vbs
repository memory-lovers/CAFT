Const ADDIN_NAME="CAFT"
Const FILE_NAME=ADDIN_NAME&".xlam"

Call Exec

Sub Exec()
    Dim objExcel
    Dim objFileSys
    Dim strAdPath
    Dim strMyPath
    Dim strAdCp
    Dim strMyCp

    Set objExcel   = CreateObject("Excel.Application")
    Set objFileSys = CreateObject("Scripting.FileSystemObject")

    'Set install path for Add-In
    strAdPath = objExcel.Application.UserLibraryPath
    strAdCp   = objFileSys.BuildPath(strAdPath, FILLE_NAME)

    'Set target Add-In location path
    strMyPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
    strMyCp   = objFileSys.BuildPath(strMyPath, FILLE_NAME)

    'Copy target Add-In
    objFileSys.CopyFile strMyCp, strAdCp

    'Set enable target Add-In for Excel
    objExcel.Workbooks.Add
    With objExcel.AddIns.Add(strAdCp,True)
                 .Installed = True
    End With
    objExcel.Quit

    Set objExcel   = Nothing
    Set objFileSys = Nothing

    MsgBox "Install CAFT is Complete!"
End Sub