Imports System.Runtime.Serialization

<Serializable>
Public NotInheritable Class GeneratedResults
    Implements ISerializable

    Public Property ElapsedTime As TimeSpan
    Public Property CreatedAssembly As Reflection.Assembly
    Public Property GeneratedExecutable As Boolean
    Public Property ExecutablePath As String
    Public Property Errors As BuildException()
    Public Property Successful As Boolean

    Public Sub GetObjectData(info As SerializationInfo, context As StreamingContext) Implements ISerializable.GetObjectData
        info.AddValue(NameOf(ElapsedTime), ElapsedTime)
        info.AddValue(NameOf(CreatedAssembly), CreatedAssembly)
        info.AddValue(NameOf(GeneratedExecutable), GeneratedExecutable)
        info.AddValue(NameOf(ExecutablePath), ExecutablePath)
        info.AddValue(NameOf(Errors), Errors)
        info.AddValue(NameOf(Successful), Successful)
    End Sub
End Class
