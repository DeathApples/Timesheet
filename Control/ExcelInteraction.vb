Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelInteraction

    Private Shared Sheet As Excel.Worksheet
    Private Shared ExcelApp As New Excel.Application

    Public Shared Subjects As New List(Of Subject)

    Public Shared PathExport As String = IO.Directory.GetCurrentDirectory & "\" & "Temp.xlsx"
    Public Shared PathImport As String = IO.Directory.GetCurrentDirectory & "\" & "Students.xlsx"


    Public Shared Sub Import(Optional path As String = "")

        If path = "" Then
            path = PathImport
        End If

        ExcelApp.Workbooks.Open(path)
        Sheet = ExcelApp.Workbooks(1).Worksheets(1)

        Dim groups = InitializeGroups()
        Subjects = InitializeSubjects(groups)

        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()

    End Sub


    Public Shared Sub Export(Optional path As String = "")

        If path = "" Then
            path = PathExport
        End If

        ExcelApp.Workbooks.Add()
        Sheet = ExcelApp.Workbooks(1).Worksheets(1)

        FillSheets()

        ExcelApp.Workbooks(1).SaveAs(path)
        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()

    End Sub


    Private Shared Function InitializeGroups() As List(Of Group)

        Dim newGroup As Group
        Dim newStudent As Student
        Dim listGroups As New List(Of Group)

        Dim rowIndex As Integer = 2

        Do While Sheet.Cells(rowIndex, 2).Value <> ""

            newGroup = New Group(Sheet.Cells(rowIndex, 2).Value)
            rowIndex += 1

            Do While Sheet.Cells(rowIndex, 2).Value <> ""

                Dim name As String = Sheet.Cells(rowIndex, 2).Value
                Dim surname As String = Sheet.Cells(rowIndex, 3).Value
                Dim patronymic As String = Sheet.Cells(rowIndex, 4).Value

                newStudent = New Student(name, surname, patronymic)
                newGroup.AddStudent(newStudent)
                rowIndex += 1

            Loop

            listGroups.Add(newGroup)
            rowIndex += 1

        Loop

        Return listGroups

    End Function


    Private Shared Function InitializeSubjects(groups As List(Of Group)) As List(Of Subject)

        Dim newSubject As Subject
        Dim listSubjects As New List(Of Subject)

        Dim rowIndex As Integer = 2

        Do While Sheet.Cells(rowIndex, 6).Value <> ""

            newSubject = New Subject(Sheet.Cells(rowIndex, 6).Value, Sheet.Cells(rowIndex, 7).Value)
            rowIndex += 1

            For Each group In groups

                If group.Name = Sheet.Cells(rowIndex, 6).Value Then

                    newSubject.AddGroup(group.Copy())

                End If

                rowIndex += 1

            Next

            listSubjects.Add(newSubject)
            rowIndex += 1

        Loop

        Return listSubjects

    End Function


    Private Shared Sub FillSheets()

        ExcelApp.Workbooks(1).Worksheets(3).Delete()
        ExcelApp.Workbooks(1).Worksheets(2).Delete()

        For Each subject In Subjects

            For Each group In subject.Groups

                ExcelApp.Workbooks(1).Sheets.Add(After:=Sheet)
                Sheet = ExcelApp.Workbooks(1).Worksheets(ExcelApp.Workbooks(1).Worksheets.Count)

                Sheet.Name = subject.ShortName & " (" & group.Name & ")"
                Sheet.Range("A1", "C1").ColumnWidth = 24

                Sheet.Cells(1, 1).Value = "Фамилия"
                Sheet.Cells(1, 2).Value = "Имя"
                Sheet.Cells(1, 3).Value = "Отчество"

                Dim listDate = group.GetAllDate()
                Dim columIndex As Integer = 4

                For Each dateTime In listDate
                    Sheet.Cells(1, columIndex) = dateTime.Date
                    columIndex += 1
                Next

                Sheet.Cells(1, 4 + listDate.Count).Value = "Итоговый балл"
                Sheet.Cells(1, 4 + listDate.Count).ColumnWidth = 14

                Dim rowIndex As Integer = 2

                For Each student In group.Students

                    Sheet.Cells(rowIndex, 2).Value = student.Name
                    Sheet.Cells(rowIndex, 1).Value = student.Surname
                    Sheet.Cells(rowIndex, 3).Value = student.Patronymic

                    Dim count As Integer = 0
                    columIndex = 4

                    Do While Sheet.Cells(1, columIndex).Value.ToString() <> "Итоговый балл" And student.CountGrades > count

                        If Sheet.Cells(1, columIndex).Value = student.Grades(count).DateTime.Date Then
                            Sheet.Cells(rowIndex, columIndex).Value = student.Grades(count).Value
                            Sheet.Cells(rowIndex, columIndex).ColumnWidth = 10
                            count += 1
                        End If

                        columIndex += 1

                    Loop

                    Sheet.Cells(rowIndex, 4 + listDate.Count).Value = student.TotalScore

                    rowIndex += 1

                Next

            Next

        Next

        ExcelApp.Workbooks(1).Worksheets(1).Delete()

    End Sub

End Class
