Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim n = TextBox1.Text

        TextBox2.Text = f(n)
    End Sub

    Function f(ByRef n2)

        If n2 = 0 Or n2 = 1 Then
            Return n2
        Else
            Return f(n2 - 1) + f(n2 - 2)
        End If

    End Function

End Class
