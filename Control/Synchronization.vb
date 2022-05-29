Public Class Synchronization

    Public Shared PathXmlFile As String = IO.Directory.GetCurrentDirectory & "\" & "Temp.xml"


    Public Shared Sub Serialize(Optional path As String = "")

        If path = "" Then
            path = PathXmlFile
        End If

        Dim serializer As New Xml.Serialization.XmlSerializer(GetType(List(Of Subject)))
        Dim file As New IO.StreamWriter(path)

        serializer.Serialize(file, ExcelInteraction.Subjects)

        file.Close()

    End Sub


    Public Shared Function Deserialize(Optional path As String = "") As Boolean

        If path = "" Then
            path = PathXmlFile
        End If

        If IO.File.Exists(path) Then

            Dim deserializer As New Xml.Serialization.XmlSerializer(GetType(List(Of Subject)))
            Dim file As New IO.StreamReader(path)

            ExcelInteraction.Subjects = CType(deserializer.Deserialize(file), List(Of Subject))

            file.Close()

            Return True

        End If

        Return False

    End Function

End Class
