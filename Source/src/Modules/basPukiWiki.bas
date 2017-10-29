Attribute VB_Name = "basPukiWiki"
Option Explicit
'-----------------------------------------------------------------------
'アイデアマンズ株式会社 Wikiサポート "Excel" アドイン を改変。
'https://www.ideamans.com/software/wiki-support-addin/
'-----------------------------------------------------------------------
'対策内容
'  Ribbon対応およびクリップボード処理の安定化
'-----------------------------------------------------------------------
'---選択範囲を書式付きWiki形式テキストに変換し、クリップボードにコピー
Sub CopyForWikiWithFormat()
    CopyForWiki True
End Sub

'---選択範囲を書式無しWiki形式テキストに変換し、クリップボードにコピー
Sub CopyForWikiWithoutFormat()
    CopyForWiki False
End Sub

'---選択範囲をWiki形式テキストに変換し、クリップボードにコピー
Private Sub CopyForWiki(Optional ByVal bWithFormat As Boolean = True)

    Dim rng As Range
    Dim r As Integer, c As Integer, i As Integer, j As Integer
    Dim matrix() As String
    Dim wiki As String, original As String, align As String
    
    '---Wikiテキスト用のバッファを確保
    ReDim matrix(Selection.Rows.Count, Selection.Columns.Count) As String
    
    '---セルを走査
    For r = 1 To Selection.Rows.Count
        For c = 1 To Selection.Columns.Count
            If matrix(r, c) <> ">" And matrix(r, c) <> "~" Then
                Set rng = Selection.Cells(r, c)
                '---値の設定
                matrix(r, c) = Replace(Trim(rng.Text), vbLf, "&br;")
                
                '---書式情報
                If bWithFormat Then
                    '---強調
                    If rng.Font.Bold Then
                        matrix(r, c) = "''" & matrix(r, c) & "''"
                    End If
                    
                    '---斜体
                    If rng.Font.Italic Then
                        matrix(r, c) = "'''" & matrix(r, c) & "'''"
                    End If
                    
                    '---取消線
                    If rng.Font.Strikethrough Then
                        matrix(r, c) = "%%" & matrix(r, c) & "%%"
                    End If
                    
                    '---文字寄せ
                    If rng.HorizontalAlignment = xlRight Or (rng.HorizontalAlignment = xlGeneral And IsNumeric(rng.Value)) Then
                        align = "RIGHT:"
                    ElseIf rng.HorizontalAlignment = xlCenter Then
                        align = "CENTER:"
                    Else
                        align = "LEFT:"
                    End If
                    matrix(r, c) = align & matrix(r, c)
                    
                    '---文字色
                    If rng.Font.ColorIndex <> xlAutomatic Then
                        matrix(r, c) = "COLOR(" & ToHexColor(rng.Font.color) & "):" & matrix(r, c)
                    End If
                    
                    '---背景色
                    If rng.Interior.Pattern = xlSolid And rng.Interior.ColorIndex <> xlAutomatic And rng.Interior.ColorIndex <> xlNone Then
                        matrix(r, c) = "BGCOLOR(" & ToHexColor(rng.Interior.color) & "):" & matrix(r, c)
                    End If
                End If
                
                '---結合セル
                original = matrix(r, c)
                If rng.MergeCells Then
                    For i = 1 To rng.MergeArea.Rows.Count
                        For j = 1 To rng.MergeArea.Columns.Count
                            If i = 1 And j = rng.MergeArea.Columns.Count Then
                                matrix(r + i - 1, c + j - 1) = original
                            ElseIf i = 1 Then
                                matrix(r + i - 1, c + j - 1) = ">"
                            Else
                                matrix(r + i - 1, c + j - 1) = "~"
                            End If
                        Next
                    Next
                    
                    '---結合セルの外側へスキップ
                    c = c + rng.MergeArea.Columns.Count - 1
                End If
                
            End If
        Next
    Next
    
    '---パイプ(|)区切りで連結
    For r = 1 To Selection.Rows.Count
        wiki = wiki & "|"
        For c = 1 To Selection.Columns.Count
            wiki = wiki & matrix(r, c) & "|"
        Next
        wiki = wiki & vbNewLine
    Next
    
    '---クリップボードにコピー
    SetClipText wiki
    
End Sub
'---カラー値をHTML式16進数の文字列に変換
Private Function ToHexColor(color As Long) As String
    Dim sColor As String
    sColor = Hex(color)
    sColor = String(6 - Len(sColor), "0") & sColor
    
    '---RとBの入れ替え(MS仕様とHTML仕様の違い)
    sColor = Right(sColor, 2) & Mid(sColor, 3, 2) & Left(sColor, 2)
    ToHexColor = "#" & sColor
End Function

