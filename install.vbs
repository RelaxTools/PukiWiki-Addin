' -------------------------------------------------------------------------------
' PukiWiki-Addin �C���X�g�[���X�N���v�g Ver.1.0.0
' -------------------------------------------------------------------------------
' �Q�l�T�C�g
' ���� SE �̂Ԃ₫
' VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' �C��
'   1.0.0 �V�K�쐬
' -------------------------------------------------------------------------------
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'�A�h�C������ݒ� 
addInName = "PukiWiki Addin" 
addInFileName = "PukiWiki.xlam"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

IF Not objFileSys.FileExists(addInFileName) THEN
   MsgBox "Zip�t�@�C����W�J���Ă�����s���Ă��������B", vbExclamation, addInName 
   WScript.Quit 
END IF

'�C���X�g�[����p�X�̍쐬 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
strPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
installPath = strPath  & addInFileName

IF MsgBox(addInName & " ���C���X�g�[�����܂����H" , vbYesNo + vbQuestion, addInName) = vbNo Then 
  WScript.Quit 
End IF

'�t�@�C���R�s�[(�㏑��) 
objFileSys.CopyFile  addInFileName ,installPath , True


Set objFileSys = Nothing

'Excel �C���X�^���X�� 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'�A�h�C���o�^ 
Set objAddin = objExcel.AddIns.Add(installPath, True) 
objAddin.Installed = True

'Excel �I�� 
objExcel.Quit
Set objAddin = Nothing 
Set objExcel = Nothing

IF Err.Number = 0 THEN 
   MsgBox "�A�h�C���̃C���X�g�[�����I�����܂����B", vbInformation, addInName 

  Set objFolder = objShell.NameSpace(strPath)
  Set objFile = objFolder.ParseName(addInFileName)
  objFile.InvokeVerb("properties")
  MsgBox "�C���^�[�l�b�g����擾�����t�@�C����Excel���u���b�N�����ꍇ������܂��B" & vbCrlf & "�v���p�e�B�E�B���h�E���J���܂��̂Łu�u���b�N�̉����v���s���Ă��������B" & vbCrLf & vbCrLf & "�v���p�e�B�Ɂu�u���b�N�̉����v���\������Ȃ��ꍇ�͓��ɑ���̕K�v�͂���܂���B", vbExclamation, addInName 

ELSE 
   MsgBox "�G���[���������܂����B" & vbCrLF & "Excel���N�����Ă���ꍇ�͏I�����Ă��������B", vbExclamation, addInName 
    WScript.Quit 
End IF

Set objWshShell = Nothing 

