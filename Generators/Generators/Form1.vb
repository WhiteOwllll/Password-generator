Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            TextBox2.Text = RandomName(TextBox1.Text, True)
        Catch ex As Exception
            MsgBox("Укажите длинну пароля", , "Generator")
        End Try
    End Sub
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Char.IsDigit(e.KeyChar) = False Then e.Handled = True
        If e.KeyChar = ChrW(Keys.Back) Then e.Handled = False
    End Sub
    Public Function RandomName(ByVal StringLength As Integer, ByVal AddNumbers As Boolean) As String
        Dim Tmp As Integer
        Dim S As Boolean
        For I As Integer = 1 To StringLength
            Do
                Randomize()
                S = True
                Tmp = Rnd() * 122
                If Tmp < 48 Then S = False
                If AddNumbers = False Then
                    If Tmp > 47 And Tmp < 58 Then S = False
                End If
                If Tmp > 57 And Tmp < 97 Then S = False
                If Tmp > 90 And Tmp < 97 Then S = False
                If S = True Then Exit Do
            Loop
            RandomName &= Chr(Tmp)
        Next
    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            If TextBox1.Text <= 6 Then Label2.Text = "Слабый пароль"
            If TextBox1.Text >= 7 And TextBox1.Text <= 10 Then Label2.Text = "Средний пароль"
            If TextBox1.Text > 10 Then Label2.Text = "Хороший пароль"
        Catch ex As Exception
            Label2.Text = ""
        End Try
        
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
      
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
