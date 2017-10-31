' -------------------------------------------------------------------------------
' PukiWiki-Addin インストールスクリプト Ver.1.0.0
' -------------------------------------------------------------------------------
' 参考サイト
' ある SE のつぶやき
' VBScript で Excel にアドインを自動でインストール/アンインストールする方法
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' 修正
'   1.0.0 新規作成
' -------------------------------------------------------------------------------
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin

'アドイン情報を設定 
addInName = "PukiWiki Addin" 
addInFileName = "PukiWiki.xlam"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

IF Not objFileSys.FileExists(addInFileName) THEN
   MsgBox "Zipファイルを展開してから実行してください。", vbExclamation, addInName 
   WScript.Quit 
END IF

'インストール先パスの作成 
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName] 
strPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\"
installPath = strPath  & addInFileName

IF MsgBox(addInName & " をインストールしますか？" , vbYesNo + vbQuestion, addInName) = vbNo Then 
  WScript.Quit 
End IF

'ファイルコピー(上書き) 
objFileSys.CopyFile  addInFileName ,installPath , True


Set objFileSys = Nothing

'Excel インスタンス化 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Workbooks.Add

'アドイン登録 
Set objAddin = objExcel.AddIns.Add(installPath, True) 
objAddin.Installed = True

'Excel 終了 
objExcel.Quit
Set objAddin = Nothing 
Set objExcel = Nothing

IF Err.Number = 0 THEN 
   MsgBox "アドインのインストールが終了しました。", vbInformation, addInName 

  Set objFolder = objShell.NameSpace(strPath)
  Set objFile = objFolder.ParseName(addInFileName)
  objFile.InvokeVerb("properties")
  MsgBox "インターネットから取得したファイルはExcelよりブロックされる場合があります。" & vbCrlf & "プロパティウィンドウを開きますので「ブロックの解除」を行ってください。" & vbCrLf & vbCrLf & "プロパティに「ブロックの解除」が表示されない場合は特に操作の必要はありません。", vbExclamation, addInName 

ELSE 
   MsgBox "エラーが発生しました。" & vbCrLF & "Excelが起動している場合は終了してください。", vbExclamation, addInName 
    WScript.Quit 
End IF

Set objWshShell = Nothing 

