Imports System.Reflection

Public Class Generator
    Implements IDisposable

    Public Shared Function IsValidName(ByVal name As String, Optional includeDot As Boolean = False) As Boolean
        For Each c As Char In name.ToCharArray
            Select Case c
                Case "a" To "z", "A" To "Z", "0" To "9", "_"
                    'ok
                Case "."
                    If includeDot Then
                        'ok
                    Else
                        Return False
                    End If
                Case Else
                    Return False
            End Select
        Next
        Return True
    End Function

    Friend Shared Function DictValuesToArray(ByVal x As IDictionary(Of String, Object)) As Object()
        If x Is Nothing Then
            Return {}
        End If
        Dim a As New List(Of Object)
        For Each K As KeyValuePair(Of String, Object) In x
            a.Add(K.Value)
        Next
        Return a.ToArray
    End Function

    Public Shared Function EvaluateExpression(ByVal language As ProgrammingLanguage, expression As String, parameters As IDictionary(Of String, Object), Optional useDefaultTemplate As Boolean = False)
        Dim g As New Generator(language)
        If useDefaultTemplate Then
            If language = ProgrammingLanguage.VisualBasic Then
                Dim src = Templates.prevb.Replace("<expression>", expression).Replace("<args>", Templates.buildVbArgs(parameters))
                g.SourceCodes = {src}
            ElseIf language = ProgrammingLanguage.VisualCSharp Then
                Dim src = Templates.precsharp.Replace("<expression>", expression).Replace("<args>", Templates.buildCsharpArgs(parameters))
                g.SourceCodes = {src}
            End If
        Else
            If language = ProgrammingLanguage.VisualBasic Then
                Dim src = Templates.preevaluatorvb.Replace("<expression>", expression).Replace("<args>", Templates.buildVbArgs(parameters))
                g.SourceCodes = {src}
            ElseIf language = ProgrammingLanguage.VisualCSharp Then
                Dim src = Templates.preevaluatorcsharp.Replace("<expression>", expression).Replace("<args>", Templates.buildCsharpArgs(parameters))
                g.SourceCodes = {src}
            End If
        End If

        Try
            Return g.InvokeStatic("ev", "Program", DictValuesToArray(parameters))
        Catch ex As BuildException
            Try
                If useDefaultTemplate Then
                    If language = ProgrammingLanguage.VisualBasic Then
                        Dim src = Templates.prevb.Replace("<expression>", expression).Replace("<args>", Templates.buildVbArgs(parameters))
                        g.SourceCodes = {src}
                    ElseIf language = ProgrammingLanguage.VisualCSharp Then
                        Dim src = Templates.precsharp.Replace("<expression>", expression).Replace("<args>", Templates.buildCsharpArgs(parameters))
                        g.SourceCodes = {src}
                    End If
                Else
                    If language = ProgrammingLanguage.VisualBasic Then
                        Dim src = Templates.precallervb.Replace("<expression>", expression).Replace("<args>", Templates.buildVbArgs(parameters))
                        g.SourceCodes = {src}
                    ElseIf language = ProgrammingLanguage.VisualCSharp Then
                        Dim src = Templates.precallercsharp.Replace("<expression>", expression).Replace("<args>", Templates.buildCsharpArgs(parameters))
                        g.SourceCodes = {src}
                    End If
                End If
                g.InvokeStatic("ev", "Program", DictValuesToArray(parameters))
                Return Nothing
            Catch ec As Exception
                Throw ex
            End Try
        End Try
    End Function

    Private Const StaticBind As BindingFlags = BindingFlags.IgnoreCase Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static
    Private Const NonStaticBind As BindingFlags = BindingFlags.IgnoreCase Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance
    Private Const FindingBind As BindingFlags = BindingFlags.IgnoreCase Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.Static

    Public Property Language As ProgrammingLanguage

    Public Property SourceCodes As String()

    Public Property References As String() =
        {"System.dll", "System.Windows.Forms.dll", "System.Drawing.dll",
        "System.Core.dll", "System.Xml.dll", "System.Xml.Linq.dll", "System.Data.dll"}

    Public Property ThrowOnFailedBuild As Boolean = True

    Public Property Entry As String = "Program"

    Public Sub New(ByVal lang As ProgrammingLanguage)
        Me.Language = lang
    End Sub

    Public Sub New(ByVal lang As ProgrammingLanguage, entryClass As String, src As String)
        Me.Language = lang
        Me.Entry = entryClass
        SourceCodes = {src}
    End Sub

    Public Sub New(ByVal lang As ProgrammingLanguage, entryClass As String, srcs As String())
        Me.Language = lang
        Me.Entry = entryClass
        Me.SourceCodes = srcs
    End Sub

    Public Sub New(ByVal lang As ProgrammingLanguage, src As String)
        Me.Language = lang
        SourceCodes = {src}
    End Sub

    Public Sub New(ByVal lang As ProgrammingLanguage, srcs As String())
        Me.Language = lang
        SourceCodes = srcs
    End Sub

    Public Function Generate() As GeneratedResults
        Return New Internal(Me).CreateAssembly(Nothing)
    End Function

    Public Function CompileToExecutable(ByVal filename As String) As GeneratedResults
        Return New Internal(Me).CreateAssembly(filename)
    End Function

    Public Function InvokeStatic(ByVal methodName As String, typename As String, constructorParameters() As Object, methodParameters() As Object)
        ' Get the constructor and create an instance of the type

        Dim generated As GeneratedResults = Generate()
        If Not generated.Successful Then
            Throw generated.Errors(0)
        End If
        Dim baseAssembly As Assembly = generated.CreatedAssembly

        Dim magicType As Type = baseAssembly.GetType(typename, True, True)

        Dim magicMethod As MethodInfo = magicType.GetMethod(methodName, StaticBind)
        Dim magicValue As Object

        Try
            magicValue = magicMethod.Invoke(Nothing, methodParameters)
        Catch ex As NullReferenceException
            Throw New Exception($"Static method '{methodName}' not found.")
        End Try

        Return magicValue
    End Function

    Public Function InvokeStatic(ByVal methodName As String, typename As String, ParamArray parameters() As Object)
        Return InvokeStatic(methodName, typename, {}, parameters)
    End Function

    Public Function InvokeStatic(ByVal methodName As String, ParamArray parameters() As Object)
        Return InvokeStatic(methodName, Entry, {}, parameters)
    End Function

    Public Function Invoke(ByVal methodName As String, typename As String, constructorParameters() As Object, methodParameters() As Object)
        ' Get the constructor and create an instance of the type

        Dim generated As GeneratedResults = Generate()
        If Not generated.Successful Then
            Throw generated.Errors(0)
        End If
        Dim baseAssembly As Assembly = generated.CreatedAssembly

        Dim magicType As Type = baseAssembly.GetType(typename, True, True)
        Dim magicConstructor As ConstructorInfo = magicType.GetConstructor(Type.EmptyTypes)
        Dim magicClassObject As Object

        Try
            magicClassObject = magicConstructor.Invoke(constructorParameters)
        Catch ex As NullReferenceException
            Throw New Exception("Type could be loaded.")
        End Try

        Dim magicMethod As MethodInfo = magicType.GetMethod(methodName, NonStaticBind)
        Dim magicValue As Object

        Try
            magicValue = magicMethod.Invoke(magicClassObject, methodParameters)
        Catch ex As NullReferenceException
            Throw New Exception($"Instance method '{methodName}' not found.")
        End Try

        Return magicValue
    End Function

    Public Function Invoke(ByVal methodName As String, typename As String, ParamArray methodParameters() As Object)
        Return Invoke(methodName, typename, {}, methodParameters)
    End Function

    Public Function Invoke(ByVal methodName As String, ParamArray methodParameters() As Object)
        Return Invoke(methodName, Entry, {}, methodParameters)
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
