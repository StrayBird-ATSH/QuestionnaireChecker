Public Class Form9
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        a8 = TextBox1.Text
        b8 = TextBox2.Text
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form8.Show()
        Me.Close()
    End Sub

    Private Sub Button2_GotFocus(sender As Object, e As EventArgs) Handles Button2.GotFocus
        a8 = TextBox1.Text
        b8 = TextBox2.Text
        Form10.Show()
        Me.Close()
    End Sub
End Class