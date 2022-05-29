Public Class Student

    Public Name As String
    Public Surname As String
    Public Patronymic As String
    Public Grades As List(Of Grade)


    Structure Grade
        Dim Value As Integer
        Dim DateTime As Date
    End Structure


    Public Sub New()

        Name = ""
        Surname = ""
        Patronymic = ""

        Grades = New List(Of Grade)

    End Sub


    Public Sub New(newSurname As String, newName As String, newPatronymic As String)

        Name = newName
        Surname = newSurname
        Patronymic = newPatronymic

        Grades = New List(Of Grade)

    End Sub


    Public ReadOnly Property CountGrades As String
        Get
            Return Grades.Count
        End Get
    End Property


    Public ReadOnly Property TotalScore As String
        Get
            Dim sum As Integer

            For Each grade In Grades
                sum += grade.Value
            Next

            Return sum
        End Get
    End Property


    Public ReadOnly Property AverageGrade As String
        Get
            Return TotalScore / Grades.Count
        End Get
    End Property


    Public Sub AddGrade(newGrade As Integer, newDate As Date)

        Dim grade As Grade
        grade.Value = newGrade
        grade.DateTime = newDate

        Grades.Add(grade)

    End Sub


    Public Function Copy() As Student

        Dim newStudent As Student = New Student(Surname, Name, Patronymic)

        For Each grade In Grades
            newStudent.AddGrade(grade.Value, grade.DateTime)
        Next

        Return newStudent

    End Function

End Class
