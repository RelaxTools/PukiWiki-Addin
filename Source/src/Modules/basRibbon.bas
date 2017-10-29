Attribute VB_Name = "basRibbon"
Option Explicit

'--------------------------------------------------------------------
'リボンより受け取ったIDをそのままマクロ名として実行するラッパー関数
'--------------------------------------------------------------------
Public Sub RibbonOnAction(control As IRibbonControl)

    Dim lngPos As Long
    Dim strBuf As String
    
    On Error GoTo e
    
    strBuf = control.ID
    
    '文字列のマクロ名を実行する。
    Application.Run strBuf

   
    Exit Sub
e:
    MsgBox Err.Description, vbOKOnly + vbCritical, "PukiWiki"
End Sub
