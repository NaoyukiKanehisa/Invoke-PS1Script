'==============================================================================================
'PowerShell�̎��s�|���V�[�Ɋ֌W�Ȃ��APowerShell�̃X�N���v�g�t�@�C��(.ps1)�����s����B
'[�g����]
'  �{VBScript�t�@�C��(.vbs)��PowerShell�̃X�N���v�g�t�@�C��(.ps1)���h���b�O&�h���b�v����B
'  �܂��́A�{VBScript�t�@�C��(.vbs)��PowerShell�̃X�N���v�g�t�@�C��(.ps1)�𓯈�f�B���N�g����
'  �z�u���A����t�@�C����(���g���q������)�ɂ���VBScript�t�@�C��(.vbs)�����s�B

'PowerShell�E�C���h�E�\��(yes/no)
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