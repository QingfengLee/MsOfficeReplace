Imports System.IO
Imports System.Security.Cryptography
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class Form1
    Public Function WordReplace(FileName As String, SearchString As String, ReplaceString As String, Optional Font As String = vbNullString, Optional Size As Single = vbNull, Optional SaveFile As String = "", Optional MatchCase As Boolean = False) As Integer
        On Error GoTo ErrorMsg '函数运行时发生遇外或错误,转向错误提示信息

        Dim wordApp As New Application
        Dim wordDoc As Document
        Dim oRange As Range



        '判断将要替换的文件是否存在
        If Dir(FileName) = "" Then
            '替换文件不存在
            MsgBox("未找到" & FileName & "文件") '提示替换文件不存在信息
            WordReplace = -2 '返回替换文件不存在的值
            Exit Function '退出函数
        End If

        wordApp.Visible = False '屏蔽WORD实例窗体
        wordDoc = wordApp.Documents.Open(FileName)
        wordDoc.Activate()

        If CheckBox1.Checked Then

            For Each oRange In wordDoc.StoryRanges
                ' We only focus on Header & Body
                With oRange.Find
                    .Text = SearchString
                    .Replacement.Text = ReplaceString
                    .Forward = True
                    .Font.Name = "楷体"
                    .Font.Size = 9
                    .Wrap = Word.WdFindWrap.wdFindContinue
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With
            Next
        Else
            For Each oRange In wordDoc.StoryRanges
                ' We only focus on Header & Body
                With oRange.Find
                    .Text = SearchString
                    .Replacement.Text = ReplaceString
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindContinue
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With
            Next
        End If



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
            Label2.Text = "正在替换 " + Path.GetFileName(fn)
            WordReplace(fn, TextBox1.Text, TextBox2.Text)
        Next
        Label2.Text = "替换完成！"
    End Sub

    Private Sub TextBox3_MouseClick(sender As Object, e As EventArgs) Handles TextBox3.MouseClick
        OpenFileDialog1.ShowDialog()
        TextBox3.Text = OpenFileDialog1.FileName + "...."
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class
