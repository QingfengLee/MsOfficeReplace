Imports System.IO
Imports System.Security.Cryptography
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Public Function WordsReplace(FileName As String, SearchString As String, ReplaceString As String, Optional Font As String = vbNullString, Optional Size As Single = vbNull, Optional SaveFile As String = "", Optional MatchCase As Boolean = False) As Integer
        On Error GoTo ErrorMsg '函数运行时发生遇外或错误,转向错误提示信息

        Dim wordApp As New Microsoft.Office.Interop.Word.Application
        Dim wordDoc As Microsoft.Office.Interop.Word.Document
        Dim oRange As Microsoft.Office.Interop.Word.Range
        '判断将要替换的文件是否存在
        If Dir(FileName) = "" Then
            '替换文件不存在
            MsgBox("未找到" & FileName & "文件") '提示替换文件不存在信息
            WordReplace = -2 '返回替换文件不存在的值
            Exit Function '退出函数
        End If

        ' 对text进行切为多行
        Dim findList = SearchString.Split(ControlChars.NewLine)
        Dim replaceList = ReplaceString.Split(ControlChars.NewLine)
        If findList.Length <> replaceList.Length Then
            MsgBox("替换的内容和被替换的内容行数不一致，检查是否有多余空行。")
            Exit Function '退出函数'
        End If
        Dim length = findList.Length


        wordApp.Visible = False '屏蔽WORD实例窗体
        wordDoc = wordApp.Documents.Open(FileName)
        wordDoc.Activate()


        For Each oRange In wordDoc.StoryRanges
            ' We only focus on Header & Body
            With oRange.Find
                For i = 0 To length - 1
                    Dim findItem = findList(i).TrimStart()
                    Dim replaceItem = replaceList(i).TrimStart()
                    .Text = findItem
                    .Replacement.Text = replaceItem
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindContinue
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                Next
            End With
        Next

        wordDoc.Save()
        wordDoc.Close() '关闭文档实例
        wordApp.Quit() '关闭WORD实例
        wordDoc = Nothing '清除文件实例
        wordApp = Nothing '清除WORD实例


        Exit Function

ErrorMsg:
        MsgBox(Err.Number & ":" & Err.Description) '提示错误信息
        WordReplace = -1 '返回错误信息值
        wordDoc.Close() '关闭文档实例
        wordApp.Quit() '关闭WORD实例
        wordDoc = Nothing '清除文件实例
        wordApp = Nothing '清除WORD实例
        Exit Function

    End Function


    Public Function ExcelsReplace(FileName As String, SearchString As String, ReplaceString As String, Optional Font As String = vbNullString, Optional Size As Single = vbNull, Optional SaveFile As String = "", Optional MatchCase As Boolean = False) As Integer
        On Error GoTo ErrorMsg '函数运行时发生遇外或错误,转向错误提示信息

        Dim oXL As Microsoft.Office.Interop.Excel.Application
        Dim oWB As Microsoft.Office.Interop.Excel.Workbook
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim oRng As Microsoft.Office.Interop.Excel.Range

        '判断将要替换的文件是否存在
        If Dir(FileName) = "" Then
            '替换文件不存在
            MsgBox("未找到" & FileName & "文件") '提示替换文件不存在信息
            WordReplace = -2 '返回替换文件不存在的值
            Exit Function '退出函数
        End If
        ' 对text进行切为多行
        Dim findList = SearchString.Split(ControlChars.NewLine)
        Dim replaceList = ReplaceString.Split(ControlChars.NewLine)
        If findList.Length <> replaceList.Length Then
            MsgBox("替换的内容和被替换的内容行数不一致，检查是否有多余空行。")
            Exit Function '退出函数'
        End If
        Dim length = findList.Length

        oXL = CreateObject("Excel.Application")
        oXL.Visible = False '屏蔽excel实例窗体

        oWB = oXL.Workbooks.Open(FileName)
        oWB.Activate()

        For Each sheet In oWB.Sheets
            For i = 0 To length - 1
                Dim findItem = findList(i).TrimStart()
                Dim replaceItem = replaceList(i).TrimStart()
                sheet.Cells.Replace(findItem, replaceItem)
            Next
        Next


        oWB.Save()
        oWB.Close() '关闭文档实例
        oXL.Quit() '关闭WORD实例
        wordDoc = Nothing '清除文件实例
        wordApp = Nothing '清除WORD实例


        Exit Function

ErrorMsg:
        MsgBox(Err.Number & ":" & Err.Description) '提示错误信息
        WordReplace = -1 '返回错误信息值
        wordDoc.Close() '关闭文档实例
        wordApp.Quit() '关闭WORD实例
        wordDoc = Nothing '清除文件实例
        wordApp = Nothing '清除WORD实例
        Exit Function

    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Then
            MsgBox("文件没选！")
            Exit Sub
        End If
        If TextBox1.Text = "" Then
            MsgBox("搜索字符串没填")
            Exit Sub
        End If
        For Each fn In OpenFileDialog1.FileNames
            If fn.EndsWith("doc") Or fn.EndsWith("docx") Then
                Label2.Text = "正在替换 " + Path.GetFileName(fn)
                WordsReplace(fn, TextBox1.Text, TextBox2.Text)
            ElseIf fn.EndsWith("xls") Or fn.EndsWith("xlsx") Then
                Label2.Text = "正在替换 " + Path.GetFileName(fn)
                ExcelsReplace(fn, TextBox1.Text, TextBox2.Text)
            Else
                MsgBox("不能处理文件" + fn)
            End If
        Next
        Label2.Text = "替换完成！"
    End Sub

    Private Sub TextBox3_MouseClick(sender As Object, e As EventArgs) Handles TextBox3.MouseClick
        OpenFileDialog1.ShowDialog()
        TextBox3.Text = OpenFileDialog1.FileName + "...."
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub
End Class
