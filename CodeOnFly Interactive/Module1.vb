Module Module1

    Dim k As CodeOnFly.Core.ProgrammingLanguage = CodeOnFly.Core.ProgrammingLanguage.VisualBasic
    Dim isOnEdit As Boolean = False
    Dim m As String = ""

    Sub Main()
        Console.WriteLine($"CodeOnFly Interactive [{My.Application.Info.Version.ToString}]")
        Console.WriteLine("Write -exit to exit, -l <1=vb 2=cs> to change the language, -r to run the syntax block.")
        Do
            Console.ForegroundColor = ConsoleColor.DarkGray
            If isOnEdit Then
                Console.Write("... ")
            Else
                Console.Write(">>> ")
            End If
            Console.ForegroundColor = ConsoleColor.Gray
            Dim expression As String = Console.ReadLine.Trim
            If expression.StartsWith("-") Then
                Select Case expression
                    Case "-exit"
                        End
                    Case "-l 1"
                        k = CodeOnFly.Core.ProgrammingLanguage.VisualBasic
                        Console.WriteLine("Language syntax changed to Visual Basic")
                    Case "-l 2"
                        k = CodeOnFly.Core.ProgrammingLanguage.VisualCSharp
                        Console.WriteLine("Language syntax changed to Visual C#")
                    Case "-r"
                        If isOnEdit Then
                            isOnEdit = False
                            Try
                                Console.WriteLine(CodeOnFly.Core.Generator.EvaluateExpression(k, m, Nothing, True))
                            Catch ex As CodeOnFly.Core.BuildException
                                If ex.ErrorID = "BC30201" Then
                                    isOnEdit = True
                                    m &= expression
                                Else
                                    Console.WriteLine($"Error [{ex.ErrorID}]: " & ex.Message)
                                End If
                            Catch ek As Exception
                                Console.WriteLine($"Error [{ek.HResult}]: " & ek.Message)
                            End Try
                            m = ""
                        Else
                            Console.WriteLine("You don't are in syntax mode.")
                        End If
                End Select
            Else
                If k = Core.ProgrammingLanguage.VisualCSharp And expression.TrimEnd.EndsWith(";") Or expression.TrimEnd.EndsWith("{") And isOnEdit = False Then
                    m &= vbNewLine & expression
                    isOnEdit = True
                    Continue Do
                End If
                If isOnEdit Then
                    m &= vbNewLine & expression
                    Continue Do
                Else
                    m = expression
                End If
                Try
                    Console.WriteLine(CodeOnFly.Core.Generator.EvaluateExpression(k, m, Nothing))
                Catch ex As CodeOnFly.Core.BuildException
                    If ex.ErrorID = "BC30201" Then
                        isOnEdit = True
                    Else
                        Console.WriteLine($"Error [{ex.ErrorID}]: " & ex.Message)
                    End If
                Catch ek As Exception
                    Console.WriteLine($"Error [{ek.HResult}]: " & ek.Message)
                End Try
            End If
        Loop
    End Sub

End Module
