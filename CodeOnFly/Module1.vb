Module Module1

    Public Sub ShowHeader()
        Console.WriteLine("Bitgate CodeOnFly .NET Compiler [version " & My.Application.Info.Version.ToString & "]")
        Console.WriteLine("Running on a " & If(Environment.Is64BitOperatingSystem, "64-bits", "32-bits") & " " & My.Computer.Info.OSFullName & "[Build " & New Version(My.Computer.Info.OSVersion).Build & "]")
        Console.WriteLine()
        Console.WriteLine("Options:")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("Syntax                Description")
        Console.WriteLine("===================== =========================================================")
        Console.WriteLine("-l <language>         Specifies the programming language to be used in the compiler.")
        Console.ForegroundColor = ConsoleColor.DarkGray
        Console.WriteLine("  -l 1                Visual Basic .NET")
        Console.WriteLine("  -l 2                Visual C-Sharp")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("-f <filepath>         Passes a code file to the compiler.")
        Console.WriteLine("-h <filepath>         Passes a reference file to the compiler.")
        Console.WriteLine("-p <filepath>         Load an CodeOnFly header file to the compiler.")
        Console.WriteLine("-e <class>            Specifies the entry class.")
        Console.WriteLine("-m <method name>      Specifies the entry method (if -out is enabled, this command will")
        Console.WriteLine("                      be ignored).")
        Console.WriteLine("-n                    Indicates that the entry method is static.")
        Console.WriteLine("-s                    Suspends the pause after the execution (only with -r).")
        Console.WriteLine("-out <filepath>       Specifies the output assembly file. Leaves it blank to don't")
        Console.WriteLine("                      compile an assembly.")
        Console.WriteLine("-r                    Run the application after compilation using the 'Program.Main' method.")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("-a <value>            Pass one or more arguments to the method in which will be started.")
        Console.WriteLine("-ver                  Shows the CodeOnFly version.")
        Console.WriteLine("-i                    Starts interactive mode (if available).")
        Console.WriteLine("-i <expression>       Evaluate an expression..")
    End Sub

#Region "Arguments variables"
    Friend _language As Core.ProgrammingLanguage
    Friend _files As New List(Of String)
    Friend _references As New List(Of String)
    Friend _out As String
    Friend _load As String
    Friend _entry As String = "Program"
    Friend _method As String = "Main"
    Friend _r As Boolean = False
    Friend _s As Boolean = False
    Friend _n As Boolean = False
    Friend _a As New List(Of String)
#End Region

    Dim v As String = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location)

    Private Function GetFileVersionInfo(ByVal filename As String) As Version
        Return Version.Parse(FileVersionInfo.GetVersionInfo(filename).FileVersion)
    End Function

    Sub Terminate(message As String, errorCode As Integer)
        Console.WriteLine($"Error [{errorCode}]: " & message)
        Environment.Exit(errorCode)
        End
    End Sub

    Sub Warn(message As String)
        Console.WriteLine("Warning: " & message)
    End Sub

    Sub Pause()
        Console.CursorTop = Console.WindowHeight - 1
        Dim z As String = "Press any key to exit. . ."
        Console.Write(z & Space(Console.WindowWidth - z.Length - 1))
        Console.ReadKey()
        End
    End Sub

    Sub Main()
        Dim commandLineArguments As String() = Environment.GetCommandLineArgs
        Dim hasArguments As Boolean = commandLineArguments.Count > 1

        If Not hasArguments Then
            ShowHeader()
        Else
            Dim args() As String = Environment.GetCommandLineArgs
            For i As Integer = 1 To args.Count - 1 Step +1
                Dim carg As String = args(i)
                Dim narg As String = Nothing
                If Not i = args.Count - 1 Then
                    Dim tmp = args(i + 1)
                    If tmp.TrimStart.StartsWith("-") Or tmp.Trim.StartsWith("/") Then
                        narg = Nothing
                    Else
                        narg = tmp
                    End If
                End If
                If carg.StartsWith("/") Then
                    carg = carg.Replace("/", "-")
                End If
                Select Case carg.Trim.ToLower
                    Case "-l"
                        If narg = Nothing Then
                            Terminate("Language expected for -l.", 1)
                        Else
                            If Not IsNumeric(narg) Then Terminate("Numeric range expected for -l.", 2)
                            Select Case narg
                                Case 1
                                    _language = Core.ProgrammingLanguage.VisualBasic
                                Case 2
                                    _language = Core.ProgrammingLanguage.VisualCSharp
                            End Select
                        End If
                    Case "-f"
                        If narg = Nothing Then
                            Terminate("Full file path expected for -f.", 1)
                        Else
                            If Not IO.File.Exists(narg) Then Terminate("File '" & narg & "' doens't exists.", 2)
                            _files.Add(narg)
                        End If
                    Case "-h"
                        If narg = Nothing Then
                            Terminate("Full file path expected for -h.", 1)
                        Else
                            If Not IO.File.Exists(narg) Then Terminate("File '" & narg & "' doens't exists.", 2)
                            _references.Add(narg)
                        End If
                    Case "-e"
                        If narg = Nothing Then
                            Terminate("Entry class name expected for -e", 1)
                        Else
                            If Not Core.Generator.IsValidName(narg) Then Terminate("Invalid Class name.", 2)
                            _entry = narg
                        End If
                    Case "-m"
                        If narg = Nothing Then
                            Terminate("Method name expected for -m", 1)
                        Else
                            If Not Core.Generator.IsValidName(narg) Then Terminate("Invalid method name.", 2)
                            _method = narg
                        End If
                    Case "-p"
                        If narg = Nothing Then
                            Terminate("Header file expected.", 1)
                        Else
                            If IO.File.Exists(narg) Then
                                _load = narg
                            Else
                                Terminate("Header file file not found.", 1)
                            End If
                        End If
                    Case "-out"
                        If narg = Nothing Then
                            Terminate("Output assembly file path expected.", 1)
                        Else
                            _out = narg
                        End If
                    Case "-r"
                        _r = True
                    Case "-s"
                        _s = True
                    Case "-n"
                        _n = True
                    Case "-a"
                        If narg = Nothing Then
                            Terminate("Argument expected for -a.", 1)
                        Else
                            _a.Add(narg)
                        End If
                    Case "-i"
                        If narg = Nothing Then
                            If IO.File.Exists(v & "\interactive.exe") Then
                                Shell(v & "\interactive.exe", AppWinStyle.NormalFocus, True)
                                End
                            Else
                                Terminate("CodeOnFly Interactive not present", -1)
                            End If
                        Else
                            Try
                                Console.WriteLine(Core.Generator.EvaluateExpression(Core.ProgrammingLanguage.VisualCSharp, narg, Nothing))
                                End
                            Catch ex As Exception
                                Terminate("Error in the formula: " & ex.Message, 21)
                            End Try
                        End If
                    Case "-ver"
                        Console.WriteLine("CodeOnFly Compiler version " & My.Application.Info.Version.ToString)
                        If IO.File.Exists(v & "\interactive.exe") Then
                            Console.WriteLine("CodeOnFly Interactive version " & GetFileVersionInfo(v & "\interactive.exe").ToString)
                        Else
                            Console.WriteLine("CodeOnFly Interactive not present")
                        End If
                        If IO.File.Exists(v & "\CodeOnFly.Core.dll") Then
                            Console.WriteLine("CodeOnFly Core version " & GetFileVersionInfo(v & "\CodeOnFly.Core.dll").ToString)
                        Else
                            Console.WriteLine("CodeOnFly Core not present")
                        End If
                        If IO.File.Exists(v & "\Serializer.dll") Then
                            Console.WriteLine("BitGate Serializer version " & GetFileVersionInfo(v & "\Serializer.dll").ToString)
                        Else
                            Console.WriteLine("BitGate Serializer not present")
                        End If
                        End
                    Case Else
                        If carg.StartsWith("-") Then
                            Terminate("Invalid argument: " & carg, 512)
                        End If
                End Select
            Next

            '---------- errors
            If _load = Nothing Then
                If _files.Count = 0 Then Terminate("No files to compile.", 4)
                If _language = Nothing Then Terminate("Programming language not selected.", 6)
                If _a.Count >= 1 And Not _out = Nothing Then Terminate("-a cannot be combined with -out.", 7)
            End If
            '---------- warnings
            If _out = Nothing Then Warn("No output executable was specified. Application will not run if the -r argument is not enabled!")
            '---------- compiler
            Dim gen As Core.Generator
            If (_load <> Nothing) Then
                gen = Header.Parser.Parse(IO.File.ReadAllText(_load))
            Else
                gen = New Core.Generator(_language)
            End If

            Dim fileArray As New List(Of String)
            For Each f As String In _files
                fileArray.Add(IO.File.ReadAllText(f))
            Next
            If _load = Nothing Then
                gen.References = _references.ToArray
                gen.Entry = _entry
                gen.SourceCodes = fileArray.ToArray
                gen.ThrowOnFailedBuild = False
            End If
            If _r Then
                Console.Clear()
                If _out = Nothing Then
                    Dim k = gen.Generate
                    If k.Successful = False Then
                        Console.WriteLine("There were build errors, aborting the execution.")
                        Console.WriteLine()

                        For i As Integer = 0 To k.Errors.Count - 1 Step +1
                            Dim err = k.Errors(i)
                            If err.IsWarning Then Continue For
                            Console.WriteLine(" - Error #" & i + 1 & " at line " & err.Line & " on file " & IO.Path.GetFileName(err.Filename) & " [" & err.ErrorID & "]: " & err.Message)
                        Next
                    Else
                        If _a.Count = 0 Then
                            Try
                                If _s Then
                                    gen.InvokeStatic(_method)
                                Else
                                    gen.Invoke(_method)
                                End If
                            Catch ex As Exception
                                Terminate(ex.Message, -2)
                            End Try
                        Else
                            Try
                                If _s Then
                                    gen.InvokeStatic(_method, _a.ToArray)
                                Else
                                    gen.Invoke(_method, _a.ToArray)
                                End If
                            Catch ex As System.Reflection.TargetParameterCountException
                                Terminate($"The number of parameters for '{_method}' invocation does not match the number expected.", 1)
                            End Try
                        End If
                        If _s = False Then Pause()
                    End If
                Else
                    Dim k = gen.CompileToExecutable(_out)
                    If k.Successful = False Then
                        Console.WriteLine("There were build errors, aborting the execution.")
                        Console.WriteLine()

                        For i As Integer = 0 To k.Errors.Count - 1 Step +1
                            Dim err = k.Errors(i)
                            If err.IsWarning Then Continue For
                            Console.WriteLine(" - Error #" & i + 1 & " at line " & err.Line & " on file " & IO.Path.GetFileName(err.Filename) & " [" & err.ErrorID & "]: " & err.Message)
                        Next
                    Else
                        Shell(_out, AppWinStyle.NormalFocus, True)
                        If _s = False Then Pause()
                    End If
                End If
            Else
                Console.WriteLine("Compile results for generated assembly:")
                Dim k = If(_out = Nothing, gen.Generate, gen.CompileToExecutable(_out))
                If k.Successful = False Then
                    Dim ce, cw As Integer
                    For Each aa In k.Errors
                        If aa.IsWarning = True Then
                            cw += 1
                        Else
                            ce += 1
                        End If
                    Next
                    ' ---------
                    Console.WriteLine("   Successfull build..............: No")
                    Console.WriteLine("   Error count....................: {0}", ce)
                    Console.WriteLine("   Warning count..................: {0}", cw)
                    Console.WriteLine("   Elapsed build time.............: {0}", ($"{k.ElapsedTime.Minutes}m{k.ElapsedTime.Seconds}s ({k.ElapsedTime.Milliseconds} ms)"))
                    Console.WriteLine("   Executable path................: {0}", If(k.ExecutablePath, "<nothing>"))
                    Console.WriteLine("   Assembly full name.............: <nothing>")
                    Console.WriteLine("Error list:")
                    For i As Integer = 0 To k.Errors.Count - 1 Step +1
                        Dim err = k.Errors(i)
                        If err.IsWarning Then Continue For
                        Console.WriteLine(" - Error #" & i + 1 & " at line " & err.Line & " on file " & IO.Path.GetFileName(err.Filename) & " [" & err.ErrorID & "]: " & err.Message)
                    Next
                    Console.WriteLine("Warning list:")
                    For i As Integer = 0 To k.Errors.Count - 1 Step +1
                        Dim err = k.Errors(i)
                        If Not err.IsWarning Then Continue For
                        Console.WriteLine(" - Warning #" & i + 1 & " at line " & err.Line & " on file " & IO.Path.GetFileName(err.Filename) & " [" & err.ErrorID & "]: " & err.Message)
                    Next
                Else
                    Dim ce, cw As Integer
                    For Each aa In k.Errors
                        If aa.IsWarning = True Then
                            cw += 1
                        Else
                            ce += 1
                        End If
                    Next
                    ' ---------
                    Console.WriteLine("   Successfull build..............: Yes")
                    Console.WriteLine("   Error count....................: {0}", ce)
                    Console.WriteLine("   Warning count..................: {0}", cw)
                    Console.WriteLine("   Elapsed build time.............: {0}", ($"{k.ElapsedTime.Minutes}m{k.ElapsedTime.Seconds}s ({k.ElapsedTime.Milliseconds} ms)"))
                    Console.WriteLine("   Executable path................: {0}", If(k.ExecutablePath, "<nothing>"))
                    Console.WriteLine("   Assembly full name.............: {0}", k.CreatedAssembly.FullName)
                    Console.WriteLine("   Entry method arg. count........: {0}", If(k.CreatedAssembly.EntryPoint IsNot Nothing, k.CreatedAssembly.EntryPoint.GetParameters.Count, 0))
                    Console.WriteLine("   Entry method name..............: {0}", If(k.CreatedAssembly.EntryPoint IsNot Nothing, k.CreatedAssembly.EntryPoint.Name, "<nothing>"))
                    Console.WriteLine("   Is fully trusted...............: {0}", k.CreatedAssembly.IsFullyTrusted)
                    Console.WriteLine("   Is dynamic.....................: {0}", k.CreatedAssembly.IsDynamic)
                    Console.WriteLine("   Image runtime version..........: {0}", k.CreatedAssembly.ImageRuntimeVersion)
                    Console.WriteLine("   Location.......................: {0}", k.CreatedAssembly.Location)
                    Console.WriteLine("   Reflection only................: {0}", k.CreatedAssembly.ReflectionOnly)
                End If
            End If
        End If
    End Sub

End Module