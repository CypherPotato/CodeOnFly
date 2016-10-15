Imports System.IO
Public Class BinarySerializer

    ''' <summary>
    ''' Writes the given object instance to a binary file.
    ''' <para>Object type (and all child types) must be decorated with the [Serializable] attribute.</para>
    ''' <para>To prevent a variable from being serialized, decorate it with the [NonSerialized] attribute; cannot be applied to properties.</para>
    ''' </summary>
    ''' <typeparam name="T">The type of object being written to the XML file.</typeparam>
    ''' <param name="filePath">The file path to write the object instance to.</param>
    ''' <param name="objectToWrite">The object instance to write to the XML file.</param>
    ''' <param name="append">If false the file will be overwritten if it already exists. If true the contents will be appended to the file.</param>
    Public Shared Sub WriteToBinaryFile(Of T)(filePath As String, objectToWrite As T, Optional append As Boolean = False)
        Using stream As Stream = File.Open(filePath, If(append, FileMode.Append, FileMode.Create))
            Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            binaryFormatter.Serialize(stream, objectToWrite)
        End Using
    End Sub

    ''' <summary>
    ''' Reads an object instance from a binary file.
    ''' </summary>
    ''' <typeparam name="T">The type of object to read from the XML.</typeparam>
    ''' <param name="filePath">The file path to read the object instance from.</param>
    ''' <returns>Returns a new instance of the object read from the binary file.</returns>
    Public Shared Function ReadFromBinaryFile(Of T)(filePath As String) As T
        Using stream As Stream = File.Open(filePath, FileMode.Open)
            Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            Return DirectCast(binaryFormatter.Deserialize(stream), T)
        End Using
    End Function

End Class
