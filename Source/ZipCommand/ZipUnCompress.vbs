'--------------------------------------------------
'ZipUnCompress
'--------------------------------------------------
'ModuleName:    ZipUnCompress.vbs
'--------------------------------------------------
'Version:       2015/07/24
'--------------------------------------------------

Option Explicit

'--------------------------------------------------
'Å°Include st.vbs
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include(".\Lib\st.vbs")
'--------------------------------------------------


Call Main

Sub Main
Do
    Dim ZipFilePath
    Dim UnCompressFolderPath
    Dim Args: Set Args = WScript.Arguments 
    If Args.Count = 2 Then
        ZipFilePath = AbsoluteFilePath(CurrentDirectory, Args(0))

        UnCompressFolderPath = AbsoluteFilePath(CurrentDirectory, Args(1))

        Call ForceCreateFolder(UnCompressFolderPath)
        Call UnZip(ZipFilePath, UnCompressFolderPath)
    Else
        Call WScript.Echo("Error:ArgsCount")
        Exit Do
    End IF

    Call WScript.echo( _
        "Finish " + WScript.ScriptName)
Loop While False
End Sub

