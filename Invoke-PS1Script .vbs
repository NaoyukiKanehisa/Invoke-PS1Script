'==============================================================================================
'PowerShellの実行ポリシーに関係なく、PowerShellのスクリプトファイル(.ps1)を実行する。
'[使い方]
'  本VBScriptファイル(.vbs)にPowerShellのスクリプトファイル(.ps1)をドラッグ&ドロップする。
'  または、本VBScriptファイル(.vbs)とPowerShellのスクリプトファイル(.ps1)を同一ディレクトリに
'  配置し、同一ファイル名(※拡張子を除く)にしてVBScriptファイル(.vbs)を実行。

'PowerShellウインドウ表示(yes/no)
displaywindow = "yes"
'==============================================================================================

Set WshShell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

If Wscript.Arguments.count = 0 Then
	ps1filename = (Fso.BuildPath(Fso.GetParentFolderName(WScript.ScriptFullName),(Fso.GetBaseName(WScript.ScriptName) & ".ps1")))
Else
	ps1filename = Wscript.Arguments(0)
End If

Command = "powershell.exe -sta -WindowStyle Normal -command " & """" & "Set-Location " & """""" & (Fso.GetParentFolderName(WScript.ScriptFullName)) & """"""";" & "(Get-Content " & """""""" & ps1filename & """""""" & ") -join """"""`r`n"""""" | Invoke-Expression"""

If displaywindow = "yes" Then
	WshShell.Run(Command),1,True
Else
	WshShell.Run(Command),0,True
End If

Set WshShell = Nothing
Set Fso = Nothing