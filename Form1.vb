Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not Synchronization.Deserialize Then
            Button2_Click(Nothing, Nothing)
        End If

        FillComboBox()

    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            If SaveFileDialog1.FileName.Split(".").Last() = "xml" Then
                Synchronization.Serialize(SaveFileDialog1.FileName)
            ElseIf SaveFileDialog1.FileName.Split(".").Last() = "xlsx" Then
                ExcelInteraction.Export(SaveFileDialog1.FileName)
            Else
                Synchronization.Serialize(SaveFileDialog1.FileName & ".xml")
            End If

        Else
            Synchronization.Serialize()
        End If

    End Sub


    Private Sub FillComboBox()

        For Each subject In ExcelInteraction.Subjects
            ComboBox1.Items.Add(subject.Name)
        Next

        ComboBox1.SelectedIndex = 0

    End Sub


    Private Sub UpdateComboBox()

        ComboBox3.Items.Clear()

        For Each student In ExcelInteraction.Subjects(ComboBox1.SelectedIndex).Groups(ComboBox2.SelectedIndex).Students
            ComboBox3.Items.Add(student.Surname & " " & student.Name.ToString().First() & ". " & student.Patronymic.ToString().First() & ".")
        Next

        ComboBox3.SelectedIndex = 0

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        ComboBox2.Items.Clear()

        For Each group In ExcelInteraction.Subjects(ComboBox1.SelectedIndex).Groups
            ComboBox2.Items.Add(group.Name)
        Next

        ComboBox2.SelectedIndex = 0

        UpdateComboBox()

    End Sub


    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        If ComboBox1.SelectedItem <> "" Then
            UpdateComboBox()
        End If

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox1.SelectedItem <> "" And ComboBox2.SelectedItem <> "" And ComboBox3.SelectedItem <> "" And TextBox1.Text <> "" Then
            ExcelInteraction.Subjects(ComboBox1.SelectedIndex).Groups(ComboBox2.SelectedIndex).Students(ComboBox3.SelectedIndex).AddGrade(CInt(TextBox1.Text), DateTimePicker1.Value)
            TextBox1.Text = ""
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.Multiselect = False

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            If OpenFileDialog1.FileName.Split(".").Last() = "xml" Then
                Synchronization.Deserialize(OpenFileDialog1.FileName)
            Else
                ExcelInteraction.Import(OpenFileDialog1.FileName)
            End If
        Else
            Application.Exit()
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then

            If SaveFileDialog1.FileName.Split(".").Last() = "docx" Then
                WordInteraction.Export(SaveFileDialog1.FileName)
            Else
                WordInteraction.Export(SaveFileDialog1.FileName & ".docx")
            End If
        End If

    End Sub

End Class
