Friend Module Templates
    Friend Function buildCsharpArgs(ByVal values As IDictionary(Of String, Object)) As String
        Dim x As String = ""
        If values Is Nothing Then
            Return ""
        End If
        For Each value As KeyValuePair(Of String, Object) In values
            For Each c As Char In value.Key
                Select Case c
                    Case "a" To "z", "A" To "Z", "_"
                    Case Else
                        Throw New InvalidExpressionException("Invalid name: " & value.Key)
                End Select
            Next
            x &= $"{value.Value.GetType.FullName} {value.Key}, "
        Next
        Return x.Substring(0, x.Length - 2)
    End Function

    Friend Function buildVbArgs(ByVal values As IDictionary(Of String, Object)) As String
        Dim x As String = ""
        If values Is Nothing Then
            Return ""
        End If
        For Each value As KeyValuePair(Of String, Object) In values
            For Each c As Char In value.Key
                Select Case c
                    Case "a" To "z", "A" To "Z", "_"
                    Case Else
                        Throw New InvalidExpressionException("Invalid name: " & value.Key)
                End Select
            Next
            x &= $"byval {value.Key} as {value.Value.GetType.FullName}, "
        Next
        Return x.Substring(0, x.Length - 2)
    End Function

    Friend prevb As String = "imports System
imports System.Windows.Forms
imports System.Drawing
imports System.Xml

class Program
    public shared function ev(<args>)
        <expression>
    end function
end class"

    Friend precallervb As String = "imports System
imports System.Windows.Forms
imports System.Drawing
imports System.Xml

class Program
    public shared sub ev(<args>)
        call <expression>
    end sub
end class"

    Friend preevaluatorvb As String = "imports System
imports System.Windows.Forms
imports System.Drawing
imports System.Xml

class Program
    public shared function ev(<args>)
        return <expression>
    end function
end class"

    Friend precsharp As String = "using System;
using System.Windows.Forms;
using System.Drawing;
using System.Xml;

class Program
{
    public static Object ev(<args>)
    {
        <expression>
    }
}"

    Friend precallercsharp As String = "using System;
using System.Windows.Forms;
using System.Drawing;
using System.Xml;

class Program
{
    public static void ev(<args>)
    {
        <expression>;
    }
}"

    Friend preevaluatorcsharp As String = "using System;
using System.Windows.Forms;
using System.Drawing;
using System.Xml;

class Program
{
    public static Object ev(<args>)
    {
        return <expression>;
    }
}"
End Module
