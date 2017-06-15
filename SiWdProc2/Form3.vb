Public Class Form3

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (TextBox1.Text.Length = 0) Then
            Return
        End If
        Dim source1 As New BindingSource
        source1.DataSource = DGV.DataSource
        source1.Filter = " freeform LIKE '%" & TextBox1.Text & "%'"
        DGV.DataSource = Nothing
        DGV.DataSource = source1.DataSource
    End Sub

    Private Sub DGV_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellDoubleClick
        Dim sz As New Size(500, 500)
        Dim pb As New WebBrowser With {.Size = sz, .Visible = True}
        'MessageBox.Show("message", "test box", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)
        Me.Controls.Add(pb)
        'pb.Visible = True
        pb.BringToFront()

    End Sub
End Class