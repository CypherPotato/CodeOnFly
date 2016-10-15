Public Class BuildException
    Inherits Exception
    Public ReadOnly Property Line As Integer
    Public ReadOnly Property IsWarning As Boolean
    Public ReadOnly Property Filename As String
    Public ReadOnly Property ErrorID As String
    Friend Sub New(ByVal err As CodeDom.Compiler.CompilerError)
        MyBase.New(err.ErrorText)
        ErrorID = err.ErrorNumber
        IsWarning = err.IsWarning
        Filename = err.FileName
        Line = err.Line
    End Sub
End Class
