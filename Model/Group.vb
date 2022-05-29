Public Class Group

    Public Name As String
    Public Students As List(Of Student)


    Public Sub New()
        Name = ""
        Students = New List(Of Student)
    End Sub


    Public Sub New(newName As String)
        Name = newName
        Students = New List(Of Student)
    End Sub


    Public ReadOnly Property CountStudents As String
        Get
            Return Students.Count
        End Get
    End Property


    Public ReadOnly Property AverageGrade As String
        Get
            Dim sum As Integer

            For Each student In Students
                sum = sum + student.TotalScore
            Next

            Return sum / Students.Count
        End Get
    End Property


    Public Sub AddStudent(newStudent As Student)
        Students.Add(newStudent)
    End Sub


    Public Function GetAllDate() As List(Of Date)

        Dim listGrades As New List(Of Date)

        For Each student In Students

            For Each grade In student.Grades

                If Not listGrades.Contains(grade.DateTime) Then

                    listGrades.Add(grade.DateTime)

                End If

            Next

        Next

        listGrades.Sort()

        Return listGrades

    End Function

    Public Function Copy() As Group

        Dim newGroup As Group = New Group(Name)

        For Each student In Students
            newGroup.AddStudent(student.Copy())
        Next

        Return newGroup

    End Function

End Class
