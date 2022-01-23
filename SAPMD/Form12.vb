Public Class Form12

    Public elSimbolo As String
    Public camPo1 As String
    Public camPo2 As String


    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Label2.Text = camPo1
        Label3.Text = camPo2

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add(">")
        ComboBox1.Items.Add("<")
        ComboBox1.SelectedIndex = 0

        Me.CenterToScreen()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        elSimbolo = ComboBox1.Items(ComboBox1.SelectedIndex)
        Me.Close()
    End Sub
End Class