
Imports System.CodeDom
Friend NotInheritable Class Internal

    Friend mg As Generator
    Public Sub New(ByVal g As Generator)
        mg = g
    End Sub

    Public Function CreateAssembly(ByVal exeName As String) As GeneratedResults

        ' If String.IsNullOrWhiteSpace(exeName) Then
        '     ' generates a temporary exe
        '     If Not IO.Directory.Exists(My.Computer.FileSystem.SpecialDirectories.Temp & "\CodeOnFlyTemp\") Then
        '         IO.Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Temp & "\CodeOnFlyTemp\")
        '     End If
        '     Dim fp = My.Computer.FileSystem.SpecialDirectories.Temp & "\CodeOnFlyTemp\" & IO.Path.GetFileName(IO.Path.GetRandomFileName) & ".exe"
        '     IO.File.Create(fp).Close()
        '     exeName = fp
        '     MsgBox(exeName)
        ' End If

        Dim x As New GeneratedResults
        Dim sp As New Stopwatch : sp.Start()

        Dim K As New Compiler.CompilerParameters
        If Not String.IsNullOrWhiteSpace(exeName) Then
            K.GenerateExecutable = True
            K.OutputAssembly = exeName
        End If
        K.GenerateInMemory = True
        K.IncludeDebugInformation = False
        K.ReferencedAssemblies.AddRange(mg.References)

        If Generator.IsValidName(mg.Entry, True) Then
            K.MainClass = mg.Entry
        Else
            Throw New Exception("Invalid name: " & mg.Entry)
        End If

        '-----------------------------------------------------------

        Dim generated As Compiler.CompilerResults = Nothing
        Select Case mg.Language
            Case ProgrammingLanguage.VisualBasic
                generated = New Microsoft.VisualBasic.VBCodeProvider().CompileAssemblyFromSource(K, mg.SourceCodes)
            Case ProgrammingLanguage.VisualCSharp
                generated = New Microsoft.CSharp.CSharpCodeProvider().CompileAssemblyFromSource(K, mg.SourceCodes)
            Case ProgrammingLanguage.VisualBasicRoslym
                generated = New Microsoft.CodeDom.Providers.DotNetCompilerPlatform.VBCodeProvider().CompileAssemblyFromSource(K, mg.SourceCodes)
            Case ProgrammingLanguage.VisualCSharpRoslym
                generated = New Microsoft.CodeDom.Providers.DotNetCompilerPlatform.CSharpCodeProvider().CompileAssemblyFromSource(K, mg.SourceCodes)
        End Select

        '-----------------------------------------------------------
        Dim hasErrors = (generated.Errors.Count <> 0)
        Dim errs As New List(Of BuildException)
        x.Successful = Not hasErrors
        If hasErrors Then
            For Each i As Compiler.CompilerError In generated.Errors
                errs.Add(New BuildException(i))
            Next
            If mg.ThrowOnFailedBuild Then
                Throw errs(0)
            End If
        End If

        '-----------------------------------------------------------

        GC.SuppressFinalize(x)

        '-----------------------------------------------------------

        If Not hasErrors Then x.CreatedAssembly = generated.CompiledAssembly
        x.ElapsedTime = sp.Elapsed
        x.Errors = errs.ToArray
        x.ExecutablePath = exeName
        x.GeneratedExecutable = Not (exeName = Nothing)

        '-----------------------------------------------------------

        sp.Stop()

        Return x
    End Function

End Class
