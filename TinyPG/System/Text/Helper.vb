Imports System

Namespace System.Text
    Public NotInheritable Class Helper
        ' Methods
        Private Sub New()
        End Sub

        Public Shared Function AddComment(ByVal comment As String) As String
            Return Helper.AddComment("//", comment)
        End Function

        Public Shared Function AddComment(ByVal commenter As String, ByVal comment As String) As String
            Return (" " & commenter & " " & comment)
        End Function

        Public Shared Function Indent(ByVal indentcount As Integer) As String
            Dim t As String = ""
            Dim i As Integer
            For i = 0 To indentcount - 1
                t = (t & "    ")
            Next i
            Return t
        End Function

        Public Shared Function Outline(ByVal text1 As String, ByVal indent1 As Integer, ByVal text2 As String, ByVal indent2 As Integer) As String
            Return ((Helper.Indent(indent1) & text1).PadRight(((indent2 * 4) Mod &H100), " "c) & text2)
        End Function

        Public Shared Function Reverse(ByVal [text] As String) As String
            Dim charArray As Char() = New Char([text].Length  - 1) {}
            Dim len As Integer = ([text].Length - 1)
            Dim i As Integer = 0
            Do While (i <= len)
                charArray(i) = [text].Chars((len - i))
                i += 1
            Loop
            Return New String(charArray)
        End Function

    End Class
End Namespace

