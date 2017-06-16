
Imports System.IO
Imports System.Windows.Documents
Imports System.Windows.Forms
Public Class Form1
    Dim db As New DC1DataContext
    Dim m As Int32
    Dim n As Int32
    Dim ttext As String
    Dim LINKfilelocation As String
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim dlg As OpenFileDialog = New OpenFileDialog
        Dim text As Integer
        Dim opfile = My.Computer.FileSystem.OpenTextFileWriter("J:\FHS\tempfile.txt", False)
        dlg.InitialDirectory = LINKfilelocation
        dlg.Title = "Open"
        dlg.Filter = "All Files (*.*)|*.*"
        If dlg.ShowDialog() = DialogResult.OK Then
            If Not (My.Computer.FileSystem.FileExists(dlg.FileName)) Then
                Return
            End If
            Dim sfile = dlg.SafeFileName

            Dim tfile As String = My.Computer.FileSystem.ReadAllText(dlg.FileName)
            If (tfile.Length < 1) Then
                Return
            End If
            text = tfile.IndexOf("J:\FHS")

            'if the string is not found
            If (text <= 0) Then
                Dim fileReader As System.IO.StreamReader
                fileReader =
                My.Computer.FileSystem.OpenTextFileReader(dlg.FileName)
                Dim stringReader As String
                Do While Not (fileReader.EndOfStream)
                    stringReader = fileReader.ReadLine()
                    ''replace returns the original string if nothing to replace
                    stringReader = stringReader.Replace("D:\XFAMILY_HISTORY\", "J:\FHS\")
                    opfile.WriteLine(stringReader)
                Loop
                opfile.Close()
                fileReader.Close()
                My.Computer.FileSystem.CopyFile("J:\FHS\tempfile.txt", "J:\FHS\LINKS\" & sfile, True)

            End If
            'if the string is found, the file has already been edited, so load it
            RichTextBox1.LoadFile(dlg.FileName, RichTextBoxStreamType.PlainText)
        End If

        TextBox2.Text = dlg.SafeFileName
        Try
            ttext = RichTextBox1.Lines(4).ToString

        Catch ex As Exception
            Return
        End Try
        ttext = ttext.Substring(4, ttext.Length - 4)
        Dim cIndex = ttext.IndexOf("</h3>")
        TextBox1.Text = ttext.Substring(0, cIndex)
    End Sub
    Private Sub check_file_exists()
        RichTextBox2.LoadFile("J:\FHS\LINKS\RTF2.txt", RichTextBoxStreamType.PlainText)
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim dlg As SaveFileDialog = New SaveFileDialog
        dlg.InitialDirectory = LINKfilelocation
        dlg.Title = "Save"
        dlg.Filter = "All Files (*.*)|*.*"
        dlg.FileName = TextBox2.Text
        If dlg.ShowDialog() = DialogResult.OK Then
            RichTextBox1.SaveFile(dlg.FileName, RichTextBoxStreamType.PlainText)
        End If
        'Dim lines(RichTextBox1.Lines.Count - 1) As String
        'RichTextBox1.Lines.CopyTo(lines, 0)
        'Dim filename As String = LINKfilelocation & TextBox2.Text
        ''IO.File.WriteAllLines(filename, lines)
        'filename = "J:\FHS\LINKS\conlinks.txt"
        'If System.IO.File.Exists(filename) = True Then
        '    Dim objWriter As New System.IO.StreamWriter(filename, True)
        '    objWriter.Write(lines)
        '    objWriter.Close()
        'Else
        '    Dim objWriter As New System.IO.StreamWriter(filename, False)
        '    objWriter.Write(lines)
        '    objWriter.Close()
        'End If   
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Dim filename = "J:\FHS\LINKS\RTF2.txt"
        'If (RichTextBox2.Lines.Count) Then
        RichTextBox2.SaveFile(filename, RichTextBoxStreamType.PlainText)
        'End If
        Me.Close()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        Dim dlg As FontDialog = New FontDialog
        dlg.Font = RichTextBox1.Font
        If dlg.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            RichTextBox1.Font = dlg.Font
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click
        Dim var1 As String
        Select Case DirectCast(sender, Button).Tag
            Case 1
                var1 = "Birth"
            Case 2
                var1 = "Marriage"
            Case 3
                var1 = "Death"
            Case 4
                var1 = "Census 1911"
            Case 5
                var1 = "Census 1901"
            Case 6
                var1 = "Census 1891"
            Case 7
                var1 = "Census 1881"
            Case 8
                var1 = "Census 1871"
            Case 9
                var1 = "Census 1861"
            Case 10
                var1 = "Census 1851"
            Case 11
                var1 = "Census 1841"
            Case 20
                var1 = "<p><a>"
            Case 21
                var1 = "<p><a href="""
            Case 22
                var1 = """>"
            Case 23
                var1 = "</p></a>"
            Case 24
                var1 = "</body>"
            Case 25
                var1 = "<p>"
            Case 26
                var1 = "</p>"
            Case Else

        End Select
        Me.RichTextBox3.Focus()
        RichTextBox3.SelectedText = var1
    End Sub


    Private Sub btnShowBrowser_Click(sender As Object, e As EventArgs) Handles btnShowBrowser.Click
        If Form2.Visible Then
            Form2.Visible = False
            Return
        End If
        If (RichTextBox1.Lines.Count = 0) Then
            Return
        End If
        Dim Image1 As Bitmap
        Dim rowIndex As Int32 = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)

        Dim test As Boolean = RichTextBox1.Lines(rowIndex).Contains("<p><a href=")
        If Not (test) Then
            MessageBox.Show("File does not exist", "", MessageBoxButtons.OK)
            Return
        End If
        Dim posInLine = RichTextBox1.SelectionStart - RichTextBox1.GetFirstCharIndexOfCurrentLine
        RichTextBox1.SelectionLength = posInLine - 12

        Dim line = RichTextBox1.Lines(rowIndex)
        Dim address = line.Substring(12, line.Length - 12)
        Dim cIndex = address.IndexOf(""">")
        address = address.Substring(0, cIndex)
        Image1 = Image.FromFile(address)
        Form2.zpb1.Image = Image1
        If My.Computer.FileSystem.FileExists(address) = False Then
            MessageBox.Show("File does not exist", "", MessageBoxButtons.OK)
            Form2.Hide()
            Return
        End If
        Form2.Show()


        'Dim bmp As New Bitmap(address)
        'Dim h = bmp.Height
        'Dim w = bmp.Width
        'If (h > 900 Or w > 1300) Then
        '    Form2.Height = 940
        '    Form2.Width = 1300
        '    Form2.PictureBox1.Height = 900
        '    Form2.PictureBox1.Width = 1300

        'Else
        '    Form2.Height = h + 40
        '    Form2.Width = w
        '    Form2.PictureBox1.Height = h
        '    Form2.PictureBox1.Width = w
        'End If
        'Form2.PictureBox1.Visible = True
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fileReader As System.IO.StreamReader
        fileReader =
        My.Computer.FileSystem.OpenTextFileReader("D:\VB_projects\SiWdProc2\SiWdProc2\init.txt")
        Dim stringReader As String
        stringReader = fileReader.ReadLine()
        If stringReader = "*LINKS FILE*" Then
            LINKfilelocation = fileReader.ReadLine()
        End If
        m = Aggregate row In db.fhdatas Into Max(row.ID)

        n = Aggregate row In db.fhdatas Into Count()

        Label1.Text = "greatest ID =" & m
        Label2.Text = n & " records"
        If (My.Computer.FileSystem.FileExists("J:\FHS\LINKS\RTF2.txt")) Then
            check_file_exists()
        End If
    End Sub

    Private Sub wrtHdr_Click(sender As Object, e As EventArgs) Handles wrtHdr.Click
        RichTextBox1.SelectedText = "<?xml version= ""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
        RichTextBox1.SelectedText = "<!--created " & Date.Now & "-->" & vbCrLf
        RichTextBox1.SelectedText = "<body>" & vbCrLf
    End Sub

    Private Sub btnWrtName_Click(sender As Object, e As EventArgs) Handles btnWrtName.Click
        If (TextBox1.Text.Length) Then
            RichTextBox1.SelectedText = "<h3>" & TextBox1.Text & "</h3>" & vbCrLf
            TextBox2.Text = TextBox1.Text & ".htm"
        Else
            MessageBox.Show("Write name in textbox", "Missing Name", MessageBoxButtons.OK)
        End If
    End Sub

    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        RichTextBox1.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim temp As String
        Dim linecount As Integer = 0
        Dim lines(RichTextBox1.Lines.Count - 1) As String
        RichTextBox1.Lines.CopyTo(lines, 0)
        RichTextBox1.Clear()

        For Each line In lines
            linecount = linecount + 1
            If (line.Contains("<a href=")) Then
                temp = line
                temp = temp.Replace("D:\", "D:\")
                RichTextBox1.AppendText(temp + vbLf)
            Else
                RichTextBox1.AppendText(line + vbLf)
            End If

        Next
    End Sub

    Private Sub WriteDBFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WriteDBFileToolStripMenuItem.Click
        Dim lines(RichTextBox1.Lines.Count - 1) As String
        Dim fd As Integer = 0
        Dim subID As String = ""
        RichTextBox1.Lines.CopyTo(lines, 0)
        Dim str As String = ""
        For Each line In lines
            If line.Contains("<body>") Then
                fd = 1
            End If
            If (fd) Then
                str = str + line + " "
            End If
        Next
        Dim item As New fhdata With
           {.ID = str.Substring(7, 5),
       .freeform = str.Substring(12, str.Length - 12)}

        Try
            Dim d = (From existingRec In db.fhdatas
                     Where existingRec.ID = item.ID
                     Select existingRec).Single()

            d.freeform = item.freeform

        Catch ex As Exception
            db.fhdatas.InsertOnSubmit(item)
        End Try

        Try
            db.SubmitChanges()
            n = db.fhdatas.Count
            MessageBox.Show("DB update/write success - " & n, "DB good", MessageBoxButtons.OK)

        Catch ex As Exception
            MessageBox.Show("DB update/write failure", "DB fail", MessageBoxButtons.OK)
        End Try

        m = Aggregate row In db.fhdatas Into Max(row.ID)

        n = Aggregate row In db.fhdatas Into Count()

        Label1.Text = "greatest ID =" & m
        Label2.Text = n & " records"
    End Sub

    Private Sub btnShowDGV_Click(sender As Object, e As EventArgs) Handles btnShowDGV.Click
        Dim db3 = New DC1DataContext
        Dim bdgsrc As New BindingSource
        Dim allrec = From recs In db3.fhdatas
                  Select recs.ID, recs.freeform
                  Order By ID Descending

        Form3.DGV.DataSource = allrec

        Dim ds As New DataSet
        ds.Tables.Add("Result")
        Dim col As DataColumn
        For Each dgvCol As DataGridViewColumn In Form3.DGV.Columns
            col = New DataColumn(dgvCol.Name)
            ds.Tables("Result").Columns.Add(col)
        Next

        Dim row As DataRow
        Dim colcount As Integer = Form3.DGV.Columns.Count - 1
        For i As Integer = 0 To Form3.DGV.Rows.Count - 1
            row = ds.Tables("Result").Rows.Add
            For Each column As DataGridViewColumn In Form3.DGV.Columns
                row.Item(column.Index) = Form3.DGV.Rows.Item(i).Cells(column.Index).Value
            Next
        Next

        Form3.DGV.DataSource = ds.Tables("Result")
        bdgsrc.DataSource = Form3.DGV.DataSource
        If (TextBox3.Text.Length > 0) Then
            Dim terms() As String = Split(TextBox3.Text)
            bdgsrc.Filter = "freeform like '%" & terms(0) & "%'"
            If terms.Length > 1 Then
                For i = 1 To terms.Length - 1
                    bdgsrc.Filter = bdgsrc.Filter & " and freeform like '%" & terms(i) & "%'"
                Next
            End If
        End If
        bdgsrc.Sort = "ID"

        Form3.DGV.DataSource = bdgsrc
        Form3.DGV.Refresh()
        Form3.DGV.Columns(0).Width = 100
        Form3.DGV.Columns(1).Width = 1200
        Form3.DGV.Sort(Form3.DGV.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        Form3.Show()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Application.Exit()
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnShowDGV_Click(sender, EventArgs.Empty)
        End If
    End Sub

    
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        RichTextBox3.Paste()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        RichTextBox3.Clear()
    End Sub


    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles btnWriteLine.Click
        Dim tlines = RichTextBox1.Lines.Count
        If (ListBox2.SelectedItem = vbNull Or tlines = 0) Then
            Return
        End If
        Dim linecount As Integer = 0
        Dim lines(tlines - 1) As String
        RichTextBox1.Lines.CopyTo(lines, 0)
        RichTextBox1.Clear()
        For Each line In lines

            linecount = linecount + 1
            If (linecount = ListBox2.SelectedItem) Then
                RichTextBox1.AppendText(RichTextBox3.Text + vbLf)
            End If
            RichTextBox1.AppendText(line + vbLf)
        Next
    End Sub

End Class