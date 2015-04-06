'--------------------------------------------------
'Standard Software Library For VBScript
'
'ModuleName:    UnCompressZip.vbs
'--------------------------------------------------
'version        2015/02/19
'--------------------------------------------------

Option Explicit

'--------------------------------------------------
'■Include Standard Software Library
'--------------------------------------------------
'FileNameには相対アドレスも指定可能
'--------------------------------------------------
'Include ".\Test\..\..\StandardSoftwareLibrary_vbs\StandardSoftwareLibrary.vbs"  
Call Include(".\Lib\StandardSoftwareLibrary.vbs")

Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    ExecuteGlobal Stream.ReadAll() 
    Call Stream.Close
End Sub
'--------------------------------------------------

Call Main

Sub Main
Do
    Dim ZipFilePath
    Dim CompressSourcePath
    Dim Args: Set Args = WScript.Arguments 
    If Args.Count = 2 Then
        CompressSourcePath = AbsoluteFilePath(CurrentDirectory, Args(0))
        ZipFilePath = AbsoluteFilePath(CurrentDirectory, Args(1))

        Call ForceCreateFolder(fso.GetParentFolderName(ZipFilePath))
        Call Zip(CompressSourcePath, ZipFilePath)
    Else
        Call WScript.Echo("Error:ArgsCount")
        Exit Do
    End IF

    Call WScript.echo( _
        "Finish " + WScript.ScriptName)
Loop While False
End Sub

