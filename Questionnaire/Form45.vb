Imports Microsoft.Office.Interop
Public Class Form45
    Public currentRow, checkRow2, checkRow1 As Short
    Public currentCell As Short
    Public whetherSkipped As Int16
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("您好，欢迎使用本应用。" + vbLf +
               "如果您比对的数据量很大，请耐心等待，在运算过程中可能出现无反应的状况，这是正常现象。" + vbLf +
               "运算速度唯一取决于您的计算机性能，其中CPU处理速度和RAM响应速度影响最大。经测试，在配备2.6GHz Intel Core i7处理器以及DDR4 RAM的计算机上进行7000次比对耗时可达到10s。另外，安装处于64bit Windows系统上的64bit Microsoft Office Excel可以额外增快处理速度。")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim newxls As Excel.Application
        Dim newbook As Excel.Workbook
        Dim sheetOne As Excel.Worksheet
        Dim sheetTwo As Excel.Worksheet
        newxls = New Excel.Application
        newbook = newxls.Workbooks.Open("D:\sheet.xlsx")
        sheetOne = newbook.Worksheets(1)
        sheetTwo = newbook.Worksheets(2)
        currentCell = 1
        currentRow = 1
        Dim contentOne, checkRowNumber2 As String
        Dim contentTwo, checkRowNumber1 As String
        checkRow1 = 1
        checkRow2 = 1
        checkRowNumber1 = sheetOne.Rows(checkRow1).cells(1).text
        checkRowNumber2 = sheetTwo.Rows(checkRow2).cells(1).text
        While (checkRowNumber1 <> "")
            checkRow1 = checkRow1 + 1
            checkRowNumber1 = sheetOne.Rows(checkRow1).cells(1).text
        End While
        While (checkRowNumber2 <> "")
            checkRow2 = checkRow2 + 1
            checkRowNumber2 = sheetTwo.Rows(checkRow2).cells(1).text
        End While
        contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
        contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
        While (contentOne = contentTwo)
            currentCell = currentCell + 1
            contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
            contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
            If (currentCell = 85) Then
                currentRow = currentRow + 1
                currentCell = 1
                contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
            End If
            If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                Exit While
            End If
        End While
        If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
            Label2.Text = "信息比对完毕，未发现不同"
        Else
            If (checkRow1 = checkRow2) Then
                Label2.Text = "发现不一致: " + vbLf + "位于第" + Str(currentRow) + "行，第" + Str(currentCell) + "列"
                If (currentCell = 1) Then
                    Form2.Show()
                ElseIf (currentCell = 2 Or currentCell = 3) Then
                    Form3.Show()
                ElseIf (currentCell = 4) Then
                    Form4.Show()
                ElseIf (currentCell = 5) Then
                    Form5.Show()
                ElseIf (currentCell = 6 Or currentCell = 84) Then
                    Form6.Show()
                ElseIf (currentCell = 7) Then
                    Form7.Show()
                ElseIf (currentCell = 8) Then
                    Form8.Show()
                ElseIf (currentCell = 9 Or currentCell = 10) Then
                    Form9.Show()
                ElseIf (currentCell = 11 Or currentCell = 12) Then
                    Form10.Show()
                ElseIf (currentCell = 13) Then
                    Form11.Show()
                ElseIf (currentCell = 14) Then
                    Form12.Show()
                ElseIf (currentCell = 15) Then
                    Form13.Show()
                ElseIf (currentCell = 16) Then
                    Form14.Show()
                ElseIf (currentCell = 17) Then
                    Form15.Show()
                ElseIf (currentCell = 18) Then
                    Form16.Show()
                ElseIf (currentCell = 19) Then
                    Form17.Show()
                ElseIf (currentCell = 20 Or currentCell = 21 Or currentCell = 22) Then
                    Form18.Show()
                ElseIf (currentCell = 25 Or currentCell = 23 Or currentCell = 24) Then
                    Form19.Show()
                ElseIf (currentCell = 28 Or currentCell = 26 Or currentCell = 27) Then
                    Form20.Show()
                ElseIf (currentCell = 31 Or currentCell = 29 Or currentCell = 30) Then
                    Form21.Show()
                ElseIf (currentCell = 34 Or currentCell = 32 Or currentCell = 33) Then
                    Form22.Show()
                ElseIf (currentCell = 37 Or currentCell = 35 Or currentCell = 36) Then
                    Form23.Show()
                ElseIf (currentCell = 40 Or currentCell = 38 Or currentCell = 39) Then
                    Form24.Show()
                ElseIf (currentCell = 43 Or currentCell = 41 Or currentCell = 42) Then
                    Form25.Show()
                ElseIf (currentCell = 46 Or currentCell = 44 Or currentCell = 45) Then
                    Form26.Show()
                ElseIf (currentCell = 47) Then
                    Form27.Show()
                ElseIf (currentCell = 50 Or currentCell = 48 Or currentCell = 49) Then
                    Form28.Show()
                ElseIf (currentCell = 53 Or currentCell = 51 Or currentCell = 52) Then
                    Form29.Show()
                ElseIf (currentCell = 56 Or currentCell = 54 Or currentCell = 55) Then
                    Form30.Show()
                ElseIf (currentCell = 58 Or currentCell = 57) Then
                    Form31.Show()
                ElseIf (currentCell = 61 Or currentCell = 59 Or currentCell = 60) Then
                    Form32.Show()
                ElseIf (currentCell = 64 Or currentCell = 62 Or currentCell = 63) Then
                    Form33.Show()
                ElseIf (currentCell = 67 Or currentCell = 65 Or currentCell = 66) Then
                    Form34.Show()
                ElseIf (currentCell = 70 Or currentCell = 68 Or currentCell = 69) Then
                    Form35.Show()
                ElseIf (currentCell = 72 Or currentCell = 71) Then
                    Form36.Show()
                ElseIf (currentCell = 73) Then
                    Form37.Show()
                ElseIf (currentCell = 75 Or currentCell = 74) Then
                    Form38.Show()
                ElseIf (currentCell = 77 Or currentCell = 76) Then
                    Form39.Show()
                ElseIf (currentCell = 79 Or currentCell = 78) Then
                    Form40.Show()
                ElseIf (currentCell = 81 Or currentCell = 80) Then
                    Form41.Show()
                ElseIf (currentCell = 83 Or currentCell = 82) Then
                    Form42.Show()
                End If
            Else
                Dim countAll, countError As Int16
                countAll = 1
                countError = 0
                While (countAll <= 14)
                    contentOne = sheetOne.Rows(currentRow).cells(countAll).text
                    contentTwo = sheetTwo.Rows(currentRow).cells(countAll).text
                    If (contentOne <> contentTwo) Then
                        countError = countError + 1
                    End If
                    countAll = countAll + 1
                End While
                If (countError > 5) Then
                    Dim rowNumberSmaller As Int16
                    If checkRow1 < checkRow2 Then
                        rowNumberSmaller = 1
                    Else
                        rowNumberSmaller = 2
                    End If
                    Label2.Text = "漏输第" + Str(rowNumberSmaller) + "组问卷编号" +
                        Str(currentRow) + vbLf + "请在手工补录后继续重新运行程序"
                    whetherSkipped = 1
                Else
                    Dim skipRow As Int16
                    skipRow = currentRow
                    currentRow = currentRow + 1
                    currentCell = 1
                    contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                    contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                    While (contentOne = contentTwo)
                        currentCell = currentCell + 1
                        contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                        contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                        If (currentCell = 85) Then
                            currentRow = currentRow + 1
                            currentCell = 1
                            contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                            contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                        End If
                        If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                            Exit While
                        End If
                    End While
                    If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                        Label2.Text = "信息比对完毕，未发现不同" + vbLf + "但是跳过了第" + Str(skipRow) + "行"
                        whetherSkipped = 1
                    Else
                        Label2.Text = "发现不一致: " + vbLf + "位于第" + Str(currentRow) +
                            "行，第" + Str(currentCell) + "列" + vbLf + "而且在之前还跳过了第" + Str(skipRow) + "行"
                        whetherSkipped = 1
                        If (currentCell = 1) Then
                            Form2.Show()
                        ElseIf (currentCell = 2 Or currentCell = 3) Then
                            Form3.Show()
                        ElseIf (currentCell = 4) Then
                            Form4.Show()
                        ElseIf (currentCell = 5) Then
                            Form5.Show()
                        ElseIf (currentCell = 6 Or currentCell = 84) Then
                            Form6.Show()
                        ElseIf (currentCell = 7) Then
                            Form7.Show()
                        ElseIf (currentCell = 8) Then
                            Form8.Show()
                        ElseIf (currentCell = 9 Or currentCell = 10) Then
                            Form9.Show()
                        ElseIf (currentCell = 11 Or currentCell = 12) Then
                            Form10.Show()
                        ElseIf (currentCell = 13) Then
                            Form11.Show()
                        ElseIf (currentCell = 14) Then
                            Form12.Show()
                        ElseIf (currentCell = 15) Then
                            Form13.Show()
                        ElseIf (currentCell = 16) Then
                            Form14.Show()
                        ElseIf (currentCell = 17) Then
                            Form15.Show()
                        ElseIf (currentCell = 18) Then
                            Form16.Show()
                        ElseIf (currentCell = 19) Then
                            Form17.Show()
                        ElseIf (currentCell = 20 Or currentCell = 21 Or currentCell = 22) Then
                            Form18.Show()
                        ElseIf (currentCell = 25 Or currentCell = 23 Or currentCell = 24) Then
                            Form19.Show()
                        ElseIf (currentCell = 28 Or currentCell = 26 Or currentCell = 27) Then
                            Form20.Show()
                        ElseIf (currentCell = 31 Or currentCell = 29 Or currentCell = 30) Then
                            Form21.Show()
                        ElseIf (currentCell = 34 Or currentCell = 32 Or currentCell = 33) Then
                            Form22.Show()
                        ElseIf (currentCell = 37 Or currentCell = 35 Or currentCell = 36) Then
                            Form23.Show()
                        ElseIf (currentCell = 40 Or currentCell = 38 Or currentCell = 39) Then
                            Form24.Show()
                        ElseIf (currentCell = 43 Or currentCell = 41 Or currentCell = 42) Then
                            Form25.Show()
                        ElseIf (currentCell = 46 Or currentCell = 44 Or currentCell = 45) Then
                            Form26.Show()
                        ElseIf (currentCell = 47) Then
                            Form27.Show()
                        ElseIf (currentCell = 50 Or currentCell = 48 Or currentCell = 49) Then
                            Form28.Show()
                        ElseIf (currentCell = 53 Or currentCell = 51 Or currentCell = 52) Then
                            Form29.Show()
                        ElseIf (currentCell = 56 Or currentCell = 54 Or currentCell = 55) Then
                            Form30.Show()
                        ElseIf (currentCell = 58 Or currentCell = 57) Then
                            Form31.Show()
                        ElseIf (currentCell = 61 Or currentCell = 59 Or currentCell = 60) Then
                            Form32.Show()
                        ElseIf (currentCell = 64 Or currentCell = 62 Or currentCell = 63) Then
                            Form33.Show()
                        ElseIf (currentCell = 67 Or currentCell = 65 Or currentCell = 66) Then
                            Form34.Show()
                        ElseIf (currentCell = 70 Or currentCell = 68 Or currentCell = 69) Then
                            Form35.Show()
                        ElseIf (currentCell = 72 Or currentCell = 71) Then
                            Form36.Show()
                        ElseIf (currentCell = 73) Then
                            Form37.Show()
                        ElseIf (currentCell = 75 Or currentCell = 74) Then
                            Form38.Show()
                        ElseIf (currentCell = 77 Or currentCell = 76) Then
                            Form39.Show()
                        ElseIf (currentCell = 79 Or currentCell = 78) Then
                            Form40.Show()
                        ElseIf (currentCell = 81 Or currentCell = 80) Then
                            Form41.Show()
                        ElseIf (currentCell = 83 Or currentCell = 82) Then
                            Form42.Show()
                        End If
                    End If
                End If
            End If
        End If
        newbook.Close()
        newxls.Quit()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim newxls As Excel.Application
        Dim newbook As Excel.Workbook
        Dim sheetOne As Excel.Worksheet
        Dim sheetTwo As Excel.Worksheet
        newxls = New Excel.Application
        newbook = newxls.Workbooks.Open("D:\sheet.xlsx")
        sheetOne = newbook.Worksheets(1)
        sheetTwo = newbook.Worksheets(2)
        currentCell = currentCell + 1
        Dim contentOne, checkRowNumber2 As String
        Dim contentTwo, checkRowNumber1 As String
        checkRowNumber1 = sheetOne.Rows(checkRow1).cells(1).text
        checkRowNumber2 = sheetTwo.Rows(checkRow2).cells(1).text
        If (whetherSkipped = 1) Then
            currentRow = currentRow + 1
        End If
        whetherSkipped = 0
        While (checkRowNumber1 <> "")
            checkRow1 = checkRow1 + 1
            checkRowNumber1 = sheetOne.Rows(checkRow1).cells(1).text
        End While
        While (checkRowNumber2 <> "")
            checkRow2 = checkRow2 + 1
            checkRowNumber2 = sheetTwo.Rows(checkRow2).cells(1).text
        End While
        contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
        contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
        While (contentOne = contentTwo)
            currentCell = currentCell + 1
            contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
            contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
            If (currentCell = 85) Then
                currentRow = currentRow + 1
                currentCell = 1
                contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
            End If
            If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                Exit While
            End If
        End While
        If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
            Label2.Text = "再次信息比对完毕，未发现不同"
        Else
            If (checkRow1 = checkRow2) Then
                Label2.Text = "发现不一致: " + vbLf + "位于第" + Str(currentRow) + "行，第" + Str(currentCell) + "列"
                If (currentCell = 1) Then
                    Form2.Show()
                ElseIf (currentCell = 2 Or currentCell = 3) Then
                    Form3.Show()
                ElseIf (currentCell = 4) Then
                    Form4.Show()
                ElseIf (currentCell = 5) Then
                    Form5.Show()
                ElseIf (currentCell = 6 Or currentCell = 84) Then
                    Form6.Show()
                ElseIf (currentCell = 7) Then
                    Form7.Show()
                ElseIf (currentCell = 8) Then
                    Form8.Show()
                ElseIf (currentCell = 9 Or currentCell = 10) Then
                    Form9.Show()
                ElseIf (currentCell = 11 Or currentCell = 12) Then
                    Form10.Show()
                ElseIf (currentCell = 13) Then
                    Form11.Show()
                ElseIf (currentCell = 14) Then
                    Form12.Show()
                ElseIf (currentCell = 15) Then
                    Form13.Show()
                ElseIf (currentCell = 16) Then
                    Form14.Show()
                ElseIf (currentCell = 17) Then
                    Form15.Show()
                ElseIf (currentCell = 18) Then
                    Form16.Show()
                ElseIf (currentCell = 19) Then
                    Form17.Show()
                ElseIf (currentCell = 20 Or currentCell = 21 Or currentCell = 22) Then
                    Form18.Show()
                ElseIf (currentCell = 25 Or currentCell = 23 Or currentCell = 24) Then
                    Form19.Show()
                ElseIf (currentCell = 28 Or currentCell = 26 Or currentCell = 27) Then
                    Form20.Show()
                ElseIf (currentCell = 31 Or currentCell = 29 Or currentCell = 30) Then
                    Form21.Show()
                ElseIf (currentCell = 34 Or currentCell = 32 Or currentCell = 33) Then
                    Form22.Show()
                ElseIf (currentCell = 37 Or currentCell = 35 Or currentCell = 36) Then
                    Form23.Show()
                ElseIf (currentCell = 40 Or currentCell = 38 Or currentCell = 39) Then
                    Form24.Show()
                ElseIf (currentCell = 43 Or currentCell = 41 Or currentCell = 42) Then
                    Form25.Show()
                ElseIf (currentCell = 46 Or currentCell = 44 Or currentCell = 45) Then
                    Form26.Show()
                ElseIf (currentCell = 47) Then
                    Form27.Show()
                ElseIf (currentCell = 50 Or currentCell = 48 Or currentCell = 49) Then
                    Form28.Show()
                ElseIf (currentCell = 53 Or currentCell = 51 Or currentCell = 52) Then
                    Form29.Show()
                ElseIf (currentCell = 56 Or currentCell = 54 Or currentCell = 55) Then
                    Form30.Show()
                ElseIf (currentCell = 58 Or currentCell = 57) Then
                    Form31.Show()
                ElseIf (currentCell = 61 Or currentCell = 59 Or currentCell = 60) Then
                    Form32.Show()
                ElseIf (currentCell = 64 Or currentCell = 62 Or currentCell = 63) Then
                    Form33.Show()
                ElseIf (currentCell = 67 Or currentCell = 65 Or currentCell = 66) Then
                    Form34.Show()
                ElseIf (currentCell = 70 Or currentCell = 68 Or currentCell = 69) Then
                    Form35.Show()
                ElseIf (currentCell = 72 Or currentCell = 71) Then
                    Form36.Show()
                ElseIf (currentCell = 73) Then
                    Form37.Show()
                ElseIf (currentCell = 75 Or currentCell = 74) Then
                    Form38.Show()
                ElseIf (currentCell = 77 Or currentCell = 76) Then
                    Form39.Show()
                ElseIf (currentCell = 79 Or currentCell = 78) Then
                    Form40.Show()
                ElseIf (currentCell = 81 Or currentCell = 80) Then
                    Form41.Show()
                ElseIf (currentCell = 83 Or currentCell = 82) Then
                    Form42.Show()
                End If
            Else
                Dim countAll, countError As Int16
                countAll = 1
                countError = 0
                While (countAll <= 14)
                    contentOne = sheetOne.Rows(currentRow).cells(countAll).text
                    contentTwo = sheetTwo.Rows(currentRow).cells(countAll).text
                    If (contentOne <> contentTwo) Then
                        countError = countError + 1
                    End If
                    countAll = countAll + 1
                End While
                If (countError > 5) Then
                    Dim rowNumberSmaller As Int16
                    If checkRow1 < checkRow2 Then
                        rowNumberSmaller = 1
                    Else
                        rowNumberSmaller = 2
                    End If
                    Label2.Text = "漏输第" + Str(rowNumberSmaller) + "组问卷编号" +
                        Str(currentRow) + vbLf + "请在手工补录后继续重新运行程序"
                Else
                    Dim skipRow As Int16
                    skipRow = currentRow
                    currentRow = currentRow + 1
                    currentCell = 1
                    contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                    contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                    While (contentOne = contentTwo)
                        currentCell = currentCell + 1
                        contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                        contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                        If (currentCell = 85) Then
                            currentRow = currentRow + 1
                            currentCell = 1
                            contentOne = sheetOne.Rows(currentRow).cells(currentCell).text
                            contentTwo = sheetTwo.Rows(currentRow).cells(currentCell).text
                        End If
                        If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                            Exit While
                        End If
                    End While
                    If (currentCell = 1 And contentOne = "" And contentTwo = "") Then
                        Label2.Text = "信息比对完毕，未发现不同" + vbLf + "但是跳过了第" + Str(skipRow) + "行"
                        whetherSkipped = 1
                    Else
                        Label2.Text = "发现不一致: " + vbLf + "位于第" + Str(currentRow) +
                            "行，第" + Str(currentCell) + "列" + vbLf + "而且在之前还跳过了第" + Str(skipRow) + "行"
                        whetherSkipped = 1
                        If (currentCell = 1) Then
                            Form2.Show()
                        ElseIf (currentCell = 2 Or currentCell = 3) Then
                            Form3.Show()
                        ElseIf (currentCell = 4) Then
                            Form4.Show()
                        ElseIf (currentCell = 5) Then
                            Form5.Show()
                        ElseIf (currentCell = 6 Or currentCell = 84) Then
                            Form6.Show()
                        ElseIf (currentCell = 7) Then
                            Form7.Show()
                        ElseIf (currentCell = 8) Then
                            Form8.Show()
                        ElseIf (currentCell = 9 Or currentCell = 10) Then
                            Form9.Show()
                        ElseIf (currentCell = 11 Or currentCell = 12) Then
                            Form10.Show()
                        ElseIf (currentCell = 13) Then
                            Form11.Show()
                        ElseIf (currentCell = 14) Then
                            Form12.Show()
                        ElseIf (currentCell = 15) Then
                            Form13.Show()
                        ElseIf (currentCell = 16) Then
                            Form14.Show()
                        ElseIf (currentCell = 17) Then
                            Form15.Show()
                        ElseIf (currentCell = 18) Then
                            Form16.Show()
                        ElseIf (currentCell = 19) Then
                            Form17.Show()
                        ElseIf (currentCell = 20 Or currentCell = 21 Or currentCell = 22) Then
                            Form18.Show()
                        ElseIf (currentCell = 25 Or currentCell = 23 Or currentCell = 24) Then
                            Form19.Show()
                        ElseIf (currentCell = 28 Or currentCell = 26 Or currentCell = 27) Then
                            Form20.Show()
                        ElseIf (currentCell = 31 Or currentCell = 29 Or currentCell = 30) Then
                            Form21.Show()
                        ElseIf (currentCell = 34 Or currentCell = 32 Or currentCell = 33) Then
                            Form22.Show()
                        ElseIf (currentCell = 37 Or currentCell = 35 Or currentCell = 36) Then
                            Form23.Show()
                        ElseIf (currentCell = 40 Or currentCell = 38 Or currentCell = 39) Then
                            Form24.Show()
                        ElseIf (currentCell = 43 Or currentCell = 41 Or currentCell = 42) Then
                            Form25.Show()
                        ElseIf (currentCell = 46 Or currentCell = 44 Or currentCell = 45) Then
                            Form26.Show()
                        ElseIf (currentCell = 47) Then
                            Form27.Show()
                        ElseIf (currentCell = 50 Or currentCell = 48 Or currentCell = 49) Then
                            Form28.Show()
                        ElseIf (currentCell = 53 Or currentCell = 51 Or currentCell = 52) Then
                            Form29.Show()
                        ElseIf (currentCell = 56 Or currentCell = 54 Or currentCell = 55) Then
                            Form30.Show()
                        ElseIf (currentCell = 58 Or currentCell = 57) Then
                            Form31.Show()
                        ElseIf (currentCell = 61 Or currentCell = 59 Or currentCell = 60) Then
                            Form32.Show()
                        ElseIf (currentCell = 64 Or currentCell = 62 Or currentCell = 63) Then
                            Form33.Show()
                        ElseIf (currentCell = 67 Or currentCell = 65 Or currentCell = 66) Then
                            Form34.Show()
                        ElseIf (currentCell = 70 Or currentCell = 68 Or currentCell = 69) Then
                            Form35.Show()
                        ElseIf (currentCell = 72 Or currentCell = 71) Then
                            Form36.Show()
                        ElseIf (currentCell = 73) Then
                            Form37.Show()
                        ElseIf (currentCell = 75 Or currentCell = 74) Then
                            Form38.Show()
                        ElseIf (currentCell = 77 Or currentCell = 76) Then
                            Form39.Show()
                        ElseIf (currentCell = 79 Or currentCell = 78) Then
                            Form40.Show()
                        ElseIf (currentCell = 81 Or currentCell = 80) Then
                            Form41.Show()
                        ElseIf (currentCell = 83 Or currentCell = 82) Then
                            Form42.Show()
                        End If
                    End If
                End If
            End If
        End If
        newbook.Close()
        newxls.Quit()
    End Sub
End Class