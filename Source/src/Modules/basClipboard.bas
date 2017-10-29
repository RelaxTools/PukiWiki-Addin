Attribute VB_Name = "basClipboard"
Option Explicit

' 32-bit Function version.
' ドライブ名からネットワークドライブを取得
#If VBA7 And Win64 Then
    'VBA7 = Excel2010以降。赤くコンパイルエラーになって見えますが問題ありません。
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
    Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function CloseClipboard Lib "user32" () As Long
    Declare Function EmptyClipboard Lib "user32" () As Long
    Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If
Public Const CF_TEXT As Long = 1  'テキストデータを読み書きする場合の定数です
'クリップボードにテキストデータを書き込むプロシージャ
Public Sub SetClipText(strData As String)

#If VBA7 And Win64 Then
  Dim lngHwnd As LongPtr, lngMEM As LongPtr
  Dim lngDataLen As LongPtr
  Dim lngRet As LongPtr
#Else
  Dim lngHwnd As Long, lngMEM As Long
  Dim lngDataLen As Long
  Dim lngRet As Long
#End If
  Dim blnErrflg As Boolean
  Const GMEM_MOVEABLE = 2

  blnErrflg = True
  
  'クリップボードをオープン
  If OpenClipboard(0&) <> 0 Then
  
    'クリップボードを空にする
    If EmptyClipboard() <> 0 Then
    
        'グローバルメモリに書き込む領域を確保してそのハンドルを取得
        lngDataLen = LenB(strData) + 1
        
        lngHwnd = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
        
        If lngHwnd <> 0 Then
      
            'グローバルメモリをロックしてそのポインタを取得
            lngMEM = GlobalLock(lngHwnd)
            
            If lngMEM <> 0 Then
        
                '書き込むテキストをグローバルメモリにコピー
                If lstrcpy(lngMEM, strData) <> 0 Then
                    'クリップボードにメモリブロックのデータを書き込み
                    lngRet = SetClipboardData(CF_TEXT, lngHwnd)
                    blnErrflg = False
                End If
                'グローバルメモリブロックのロックを解除
                lngRet = GlobalUnlock(lngHwnd)
            End If
        End If
    End If
    'クリップボードをクローズ(これはWindowsに制御が
    '戻らないうちにできる限り速やかに行う)
    lngRet = CloseClipboard()
  End If

  If blnErrflg Then MsgBox "クリップボードに情報が書き込めません", vbOKOnly, ""

End Sub

