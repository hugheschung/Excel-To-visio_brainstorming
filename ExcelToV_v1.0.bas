'v1.0 須要手動轉換編碼為uft-8格式。

Sub ExcelToV_v1.0()

' 宣告區
Dim OutputFilePath As String
Dim Content As String
Dim Headxml As String

Dim fDialog As FileDialog


' 建立選擇目錄的對話方塊
Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

If fDialog.Show = -1 Then
  ' 顯示選擇的目錄
  'MsgBox fDialog.SelectedItems(1)  debug使用
End If



' 文字檔案位置
OutputFilePath = fDialog.SelectedItems(1) + "\ExcelToV.xml"


' 開啟 OutputFilePath 文字檔，使用編號 #1 檔案代碼
Open OutputFilePath For Output As #1

' 要寫入的內容：寫入xml標頭，另外 Chr(13) & Chr(10) = \n\r
Headxml = "<?xml version=""1.0"" encoding=""UTF-8""?> " & Chr(13) & Chr(10) & "<bs:Brainstorm xmlns:bs=""http://schemas.microsoft.com/visio/2003/brainstorming"">"

' 將 Headxml 的內容寫入編號 #1 的檔案
Print #1, Headxml


For x = 1 To ActiveSheet.[A65536].End(xlUp).Row
    Call condic(x)
    
Next x


'Print #1, Topic_convert(Cells(1, 1).Value)
'Print #1, text_convert("主題1")


' 關閉編號 #1 檔案
Close #1


'MsgBox ActiveSheet.[A65536].End(xlUp).Row  '找出有多少使用的欄位
MsgBox "執行完成"

End Sub

'自定義函數區

'轉換topic的 xml標籤
Function Topic_convert(Topic As String) As String

    Content_Topic = "<bs:topic bs:TopicID=" & """" & Topic & """" & ">"
    Topic_convert = Content_Topic

End Function

'轉換text的 xml標籤
Function text_convert(text As String) As String

    Content_text = "<bs:text>" & text & "</bs:text>"
    text_convert = Content_text
  
End Function

Sub Print_end_topic(i As Integer)
    For a = 1 To i
    Print #1, "</bs:topic>"
    Next a
End Sub


Sub condic(i)
    If Len(Cells(i + 1, 1).Value) = Len(Cells(i, 1).Value) Then '遇到同階層，就加入</bs:topic>
        Print #1, Topic_convert(Cells(i, 1).Value) 'A欄 階層
        Print #1, text_convert(Cells(i, 2).Value)  'B欄 text
         Call Print_end_topic(1)
'        MsgBox "1" '測試輸出
    ElseIf Cells(i + 1, 1).Value = "" Then                      '遇到空格加入</bs:topic>1層x1,2層x2,3層x3 </bs:topic>
        Print #1, Topic_convert(Cells(i, 1).Value) 'A欄 階層
        Print #1, text_convert(Cells(i, 2).Value)  'B欄 text
        Call Print_end_topic(Len(Cells(i, 1).Value) / 2)
        Print #1, "</bs:Brainstorm>"
'        MsgBox "2" '測試輸出
    ElseIf Len(Cells(i, 1).Value) - Len(Cells(i + 1, 1).Value) > 1 Then             '遇到上層或上上層等...加入</bs:topic>
        Print #1, Topic_convert(Cells(i, 1).Value) 'A欄 階層
        Print #1, text_convert(Cells(i, 2).Value)  'B欄 text
        Call Print_end_topic((Len(Cells(i, 1).Value) - Len(Cells(i + 1, 1).Value)) / 2 + 1)
   
    Else                                                       '遇到下一層
        Print #1, Topic_convert(Cells(i, 1).Value) 'A欄 階層
        Print #1, text_convert(Cells(i, 2).Value)  'B欄 text
    End If
End Sub
