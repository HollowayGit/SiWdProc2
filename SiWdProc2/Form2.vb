Public Class Form2
    'Private originalSize As Size = Nothing
    'Private xscale As Single = 1
    'Private scaleDelta As Single = 0.00005
    'Private ratWidth, ratHeight As Double
    Dim m_PanStartPoint As Point
    Private Sub Form_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        'If e.Delta <> 0 Then
        '    If e.Delta <= 0 Then
        '        If PictureBox1.Width < 500 Then Exit Sub 'minimum 500?
        '    Else
        '        If PictureBox1.Width > 2000 Then Exit Sub 'maximum 2000?
        '    End If

        '    PictureBox1.Width += CInt(PictureBox1.Width * e.Delta / 1000)
        '    PictureBox1.Height += CInt(PictureBox1.Height * e.Delta / 1000)
        'End If

        'scaleDelta = Math.Sqrt(PictureBox1.Width * PictureBox1.Height) * 0.00005
        'If e.Delta < 0 Then
        '    xscale -= scaleDelta
        'ElseIf e.Delta >0 Then
        '    xscale += scaleDelta
        'End If

        'If e.Delta <> 0 Then
        '    PictureBox1.Size = New Size(CInt(Math.Round((originalSize.Width * ratWidth) * xscale)),
        '    CInt(Math.Round((originalSize.Height * ratHeight) * xscale)))
        'End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.AutoScroll = True
        PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        m_PanStartPoint = New Point(e.X, e.Y)
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim deltaX As Integer = (m_PanStartPoint.X - e.X)
            Dim deltaY As Integer = (m_PanStartPoint.Y - e.Y)
            Panel1.AutoScrollPosition =
                New Drawing.Point((deltaX - Panel1.AutoScrollPosition.X),
                                  (deltaY - Panel1.AutoScrollPosition.Y))
        End If
    End Sub
End Class
