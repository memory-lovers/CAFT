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
    strAdCp   = objFileSys.BuildPath(strAdPath, FILE_NAME)

    'Set disable target Add-In for Excel
    objExcel.Workbooks.Add
    With objExcel.AddIns(ADDIN_NAME)
                 .Installed = False
    End With
    objExcel.Quit

    'Delete Add-In in intall path
    objFileSys.DeleteFile strAdCp

    Set objExcel   = Nothing
    Set objFileSys = Nothing

    MsgBox "Uninstall CAFT is Complete!"
End Sub