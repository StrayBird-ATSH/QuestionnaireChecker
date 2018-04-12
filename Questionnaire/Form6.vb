Public Class Form6
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form5.Show()
        Me.Close()
    End Sub

    Private Sub Button2_GotFocus(sender As Object, e As EventArgs) Handles Button2.GotFocus
        sisters = TextBox2.Text
        a5 = TextBox1.Text
        Form7.Show()
        Me.Close()
    End Sub
End Class