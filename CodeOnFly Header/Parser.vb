Imports Core.CodeOnFly.Core
Public Class Parser
    Public Class ParserException
        Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class
    Public Shared Function Parse(ByVal str As String) As Generator
        Dim g As New Generator(ProgrammingLanguage.VisualBasic)
        Dim srcs As New List(Of String)
        Dim ref As New List(Of String)
        ' ------------------------------------------------------------------------------------------------
        Dim Lines As String() = str.Split(vbNewLine)
        For i As Integer = 0 To Lines.Count - 1 Step +1
            Dim Current As String = Lines(i).Trim
            If String.IsNullOrWhiteSpace(Current) Or Current.StartsWith(";") Then
                Continue For
            End If
            Dim Args As String() = Current.Split(" ")
            Dim First As String
            If Args.Count = 0 Then
                First = Current
            Else
                First = Args(0)
            End If
            '----------------------------------------------------------------------
            Select Case First
                Case "lang"
                    Dim Arg As String = Current.Substring(Len("lang")).Trim
                    Select Case Arg
                        Case "1", "vb"
                            g.Language = ProgrammingLanguage.VisualBasic
                        Case "2", "cs"
                            g.Language = ProgrammingLanguage.VisualCSharp
                        Case Else
                            Throw New ParserException("Invalid parser argument for 'lang' at line " & i + 1)
                    End Select

                Case "file"
                    Dim Arg As String = Current.Substring(Len("file")).Trim
                    If Not IO.File.Exists(Arg) Then
                        Throw New ParserException("Input file not found at line " & i + 1)
                    Else
                        srcs.Add(IO.File.ReadAllText(Arg))
                    End If

                Case "import"
                    Dim Arg As String = Current.Substring(Len("import")).Trim
                    If Not IO.File.Exists(Arg) Then
                        Throw New ParserException("Reference file not found at line " & i + 1)
                    Else
                        ref.Add(Arg)
                    End If

                Case "entry"
                    Dim Arg As String = Current.Substring(Len("entry")).Trim
                    g.Entry = Arg

                Case Else
                    Throw New ParserException("Invalid command '" & First & "' at line " & i + 1)
            End Select
        Next
        g.SourceCodes = srcs.ToArray
        Return g
    End Function
End Class
