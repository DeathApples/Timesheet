Public Class Subject

    Public Name As String
    Public ShortName As String
    Public Groups As List(Of Group)


    Public Sub New()
        Name = ""
        ShortName = ""
        Groups = New List(Of Group)
    End Sub


    Public Sub New(newName As String, newShortName As String)
        Name = newName
        ShortName = newShortName
        Groups = New List(Of Group)
    End Sub


    Public ReadOnly Property CountGroups As String
        Get
            Return Groups.Count
        End Get
    End Property


    Public Sub AddGroup(group As Group)
        Groups.Add(group)
    End Sub

End Class
