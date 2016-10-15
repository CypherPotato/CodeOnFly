Imports System.Deployment
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Text
Imports Microsoft.Win32

Module Module1

    ReadOnly FilesToInstall As String() = {"interactive.exe", "CodeOnFly.Core.dll", "Serializer.dll", "CodeOnFly.exe", "Setup.exe", "CodeOnFly.Header.dll"}
    ReadOnly WhereToInstall As String = "C:\CodeOnFly\"

    Sub WriteOnCenter(ByVal txt As String)
        Console.WriteLine(Space(Console.WindowWidth / 2 - (txt.Length / 2)) & txt)
    End Sub

    Sub Main()
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.Clear()
        Console.WriteLine()
        WriteOnCenter("######################################################")
        WriteOnCenter("#                 CodeOnFly Installer                #")
        WriteOnCenter("#                     version 1.0                    #")
        WriteOnCenter("#                                                    #")
        WriteOnCenter("#                      by BitGate                    #")
        WriteOnCenter("######################################################")
        Console.WriteLine()
        Console.WriteLine("CodeOnFly Installer")
        Console.WriteLine("This wizards installs CodeOnFly in this computer, so you can launch")
        Console.WriteLine("from CMD, and you can use the CodeOnFly API.")
        Console.WriteLine()
        Console.WriteLine("All files will be installed in '{0}'", WhereToInstall)
        Console.WriteLine()
        Console.WriteLine("During the CodeOnFly use, you are accepting the license in")
        Console.WriteLine("http://bitgatedev.cf/codeonfly-net")
        Console.WriteLine()
        Console.WriteLine(New String("=", Console.WindowWidth))
        Console.ForegroundColor = ConsoleColor.White
        ' ----------------------------------------------------------

        If Not UacHelper.IsProcessElevated Then
            Console.WriteLine("[!] Installation aborted.")
            Console.WriteLine("Start this application in elevated mode to install the aplication.")
            Console.WriteLine()
            Console.WriteLine("Press any key to exit. . .")
            Console.ReadKey()
            End
        Else
            Console.WriteLine("Press any key to install. . .")
            Console.ReadKey()
        End If

        ' ----------------------------------------------------------

        Console.Clear()
        If Not IO.Directory.Exists(WhereToInstall) Then
            IO.Directory.CreateDirectory(WhereToInstall)
        End If
        For i As Integer = 0 To FilesToInstall.Count - 1
            Dim file As String = IO.Directory.GetCurrentDirectory & "\" & FilesToInstall(i)
            If Not IO.File.Exists(file) Then
                Console.WriteLine("Error in the installation: File '" & IO.Path.GetFileName(file) & "' is not present.")
                Console.WriteLine("Press any key to exit. . .")
                Console.ReadKey()
                End
            Else
                Console.WriteLine("Copying files... " & Math.Round((i / FilesToInstall.Count) * 100) & "%")
                IO.File.Copy(file, WhereToInstall & IO.Path.GetFileName(file), True)
            End If
            Threading.Thread.Sleep(New Random().Next(250, 1800))
        Next

        ' ----------------------------------------------------------

        Console.WriteLine("Creating path variables...")
        Process.Start("cmd", $"/c setx PATH ""{WhereToInstall};%path%;"" -m")
        Console.WriteLine("Creating associations...")
        FileAssociation.CreateFileAssociation(".vbr", "VisualBasicCodeOnFlyScript", "Visual Basic CodeOnFly auto-run script", WhereToInstall & "CodeOnFly.exe", " -l 1 -f ""%1"" -r")
        FileAssociation.CreateFileAssociation(".csr", "CSharpCodeOnFlyScript", "C# CodeOnFly auto-run script", WhereToInstall & "CodeOnFly.exe", " -l 2 -f ""%1"" -r")

        ' ----------------------------------------------------------

        Console.Clear()
        Beep()
        Console.WriteLine("CodeOnFly was succesfully installed in this Computer. To uninstall it,")
        Console.WriteLine("simply delete the directory '{0}' from your system.", WhereToInstall)
        Console.WriteLine()
        Console.WriteLine("Press any key to exit setup. . .")
        Console.ReadKey()
    End Sub
End Module

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================


Public NotInheritable Class FileAssociation
    <System.Runtime.InteropServices.DllImport("shell32.dll")> Shared Sub _
    SHChangeNotify(ByVal wEventId As Integer, ByVal uFlags As Integer,
    ByVal dwItem1 As Integer, ByVal dwItem2 As Integer)
    End Sub


    ' Create the new file association
    '
    ' Extension is the extension to be registered (eg ".cad"
    ' ClassName is the name of the associated class (eg "CADDoc")
    ' Description is the textual description (eg "CAD Document"
    ' ExeProgram is the app that manages that extension (eg "c:\Cad\MyCad.exe")

    Shared Function CreateFileAssociation(ByVal extension As String,
    ByVal className As String, ByVal description As String,
    ByVal exeProgram As String,
    ByVal commandline As String) As Boolean
        Const SHCNE_ASSOCCHANGED = &H8000000
        Const SHCNF_IDLIST = 0

        ' ensure that there is a leading dot
        If extension.Substring(0, 1) <> "." Then
            extension = "." & extension
        End If

        Dim key1, key2, key3 As Microsoft.Win32.RegistryKey
        Try
            ' create a value for this key that contains the classname
            key1 = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(extension)
            key1.SetValue("", className)
            ' create a new key for the Class name
            key2 = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(className)
            key2.SetValue("", description)
            ' associate the program to open the files with this extension
            key3 = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(className &
            "\Shell\Open\Command")
            key3.SetValue("", exeProgram & commandline)
        Catch e As Exception
            Return False
        Finally
            If Not key1 Is Nothing Then key1.Close()
            If Not key2 Is Nothing Then key2.Close()
            If Not key3 Is Nothing Then key3.Close()
        End Try

        ' notify Windows that file associations have changed
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)
        Return True
    End Function
End Class

Public NotInheritable Class UacHelper
    Private Sub New()
    End Sub
    Private Const uacRegistryKey As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    Private Const uacRegistryValue As String = "EnableLUA"

    Private Shared STANDARD_RIGHTS_READ As UInteger = &H20000
    Private Shared TOKEN_QUERY As UInteger = &H8
    Private Shared TOKEN_READ As UInteger = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)

    <DllImport("advapi32.dll", SetLastError:=True)>
    Private Shared Function OpenProcessToken(ProcessHandle As IntPtr, DesiredAccess As UInt32, ByRef TokenHandle As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Public Shared Function GetTokenInformation(TokenHandle As IntPtr, TokenInformationClass As TOKEN_INFORMATION_CLASS, TokenInformation As IntPtr, TokenInformationLength As UInteger, ByRef ReturnLength As UInteger) As Boolean
    End Function

    Public Enum TOKEN_INFORMATION_CLASS
        TokenUser = 1
        TokenGroups
        TokenPrivileges
        TokenOwner
        TokenPrimaryGroup
        TokenDefaultDacl
        TokenSource
        TokenType
        TokenImpersonationLevel
        TokenStatistics
        TokenRestrictedSids
        TokenSessionId
        TokenGroupsAndPrivileges
        TokenSessionReference
        TokenSandBoxInert
        TokenAuditPolicy
        TokenOrigin
        TokenElevationType
        TokenLinkedToken
        TokenElevation
        TokenHasRestrictions
        TokenAccessInformation
        TokenVirtualizationAllowed
        TokenVirtualizationEnabled
        TokenIntegrityLevel
        TokenUIAccess
        TokenMandatoryPolicy
        TokenLogonSid
        MaxTokenInfoClass
    End Enum

    Public Enum TOKEN_ELEVATION_TYPE
        TokenElevationTypeDefault = 1
        TokenElevationTypeFull
        TokenElevationTypeLimited
    End Enum

    Public Shared ReadOnly Property IsUacEnabled() As Boolean
        Get
            Dim uacKey As RegistryKey = Registry.LocalMachine.OpenSubKey(uacRegistryKey, False)
            Dim result As Boolean = uacKey.GetValue(uacRegistryValue).Equals(1)
            Return result
        End Get
    End Property

    Public Shared ReadOnly Property IsProcessElevated() As Boolean
        Get
            If IsUacEnabled Then
                Dim tokenHandle As IntPtr
                If Not OpenProcessToken(Process.GetCurrentProcess().Handle, TOKEN_READ, tokenHandle) Then
                    Throw New ApplicationException("Could not get process token.  Win32 Error Code: " + Marshal.GetLastWin32Error())
                End If

                Dim elevationResult As TOKEN_ELEVATION_TYPE = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault

                Dim elevationResultSize As Integer = Marshal.SizeOf(CInt(elevationResult))
                Dim returnedSize As UInteger = 0
                Dim elevationTypePtr As IntPtr = Marshal.AllocHGlobal(elevationResultSize)

                Dim success As Boolean = GetTokenInformation(tokenHandle, TOKEN_INFORMATION_CLASS.TokenElevationType, elevationTypePtr, CUInt(elevationResultSize), returnedSize)
                If success Then
                    elevationResult = CType(Marshal.ReadInt32(elevationTypePtr), TOKEN_ELEVATION_TYPE)
                    Dim isProcessAdmin As Boolean = elevationResult = TOKEN_ELEVATION_TYPE.TokenElevationTypeFull
                    Return isProcessAdmin
                Else
                    Throw New ApplicationException("Unable to determine the current elevation.")
                End If
            Else
                Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
                Dim principal As New WindowsPrincipal(identity)
                Dim result As Boolean = principal.IsInRole(WindowsBuiltInRole.Administrator)
                Return result
            End If
        End Get
    End Property
End Class

