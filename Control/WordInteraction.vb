Imports Word = Microsoft.Office.Interop.Word

Public Class WordInteraction

    Private Shared Document As Word.Document
    Private Shared WordApp As New Word.Application

    Public Shared PathTemplate As String = IO.Directory.GetCurrentDirectory & "\" & "Template.docx"


    Public Shared Sub Export(path As String)

        For Each subject In ExcelInteraction.Subjects

            For Each group In subject.Groups

                Document = WordApp.Documents.Open(PathTemplate)
                WordApp.Visible = False

                Document.Content.Paragraphs(3).Range.Text = Document.Content.Paragraphs(3).Range.Text.Replace("{Group}", group.Name)
                Document.Content.Paragraphs(2).Range.Text = Document.Content.Paragraphs(2).Range.Text.Replace("{Subject}", subject.Name)
                Document.Content.Paragraphs(5).Range.Text = Document.Content.Paragraphs(5).Range.Text.Replace("{AverageGrade}", group.AverageGrade)
                Document.Content.Paragraphs(4).Range.Text = Document.Content.Paragraphs(4).Range.Text.Replace("{StudentCount}", group.CountStudents)

                Dim table = Document.Tables.Add(Document.Paragraphs(8).Range, group.CountStudents + 1, 4)
                table.Borders.Enable = True

                Dim rowIndex As Integer = 2

                table.Cell(1, 1).Range.Text = "Фамилия"
                table.Cell(1, 1).Range.Font.Bold = True
                table.Cell(1, 1).Range.Font.Size = 16

                table.Cell(1, 2).Range.Text = "Имя"
                table.Cell(1, 2).Range.Font.Bold = True
                table.Cell(1, 2).Range.Font.Size = 16

                table.Cell(1, 3).Range.Text = "Отчество"
                table.Cell(1, 3).Range.Font.Bold = True
                table.Cell(1, 3).Range.Font.Size = 16

                table.Cell(1, 4).Range.Text = "Итоговый балл"
                table.Cell(1, 4).Range.Font.Bold = True
                table.Cell(1, 4).Range.Font.Size = 16

                For Each student In group.Students

                    table.Cell(rowIndex, 1).Range.Text = student.Surname
                    table.Cell(rowIndex, 1).Range.Font.Bold = False
                    table.Cell(rowIndex, 1).Range.Font.Size = 16

                    table.Cell(rowIndex, 2).Range.Text = student.Name
                    table.Cell(rowIndex, 2).Range.Font.Bold = False
                    table.Cell(rowIndex, 2).Range.Font.Size = 16

                    table.Cell(rowIndex, 3).Range.Text = student.Patronymic
                    table.Cell(rowIndex, 3).Range.Font.Bold = False
                    table.Cell(rowIndex, 3).Range.Font.Size = 16

                    table.Cell(rowIndex, 4).Range.Text = student.TotalScore
                    table.Cell(rowIndex, 4).Range.Font.Bold = False
                    table.Cell(rowIndex, 4).Range.Font.Size = 16

                    rowIndex += 1

                Next

                Document.SaveAs2(IO.Directory.GetCurrentDirectory & "\" & subject.ShortName & "_" & group.Name & ".docx")
                WordApp.Documents.Close()

            Next

        Next

        WordApp.Quit()

    End Sub

End Class
