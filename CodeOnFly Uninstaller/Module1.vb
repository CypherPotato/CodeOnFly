
Module Module1

    Dim myDir = IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location) & "\"

    Dim myExe = Reflection.Assembly.GetExecutingAssembly.Location

    Sub WriteOnCenter(ByVal txt As String)
        Console.WriteLine(Space(Console.WindowWidth / 2 - (txt.Length / 2)) & txt)
    End Sub

    Sub DeleteItself()
        Process.Start("cmd.exe", "/C choice /C Y /N /D Y /T 3 & Del " & myExe)
        End
    End Sub

    Sub Uninstall()
        Console.Clear()
        For Each F As String In IO.Directory.GetFiles(CodeOnFly_Setup.WhereToInstall, "*", IO.SearchOption.AllDirectories)

        Next
    End Sub

    Sub Main()
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.Clear()
        Dim Code As Integer = New Random().Next(1000, 9999)
        Console.WriteLine()
        Console.WriteLine("Are you sure wantn to uninstall CodeOnFly from your computer?")
        Console.WriteLine("If you want, write the following number and hit enter, elsewise,")
        Console.WriteLine("close this window.")
retry:
        Console.WriteLine()
        Console.Write("   Type '" & Code & "' to confirm your decision: ")
        Dim r As String = Console.ReadLine
        If r.Trim = Code Then
            Uninstall()
        Else
            Console.WriteLine("Wrong code.")
            GoTo retry
        End If
    End Sub

End Module
