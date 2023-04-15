Imports System
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler

Namespace TinyPG.CodeGenerators.VBNet
    Public Class ScannerGenerator
        Implements ICodeGenerator
        ' Methods
        Friend Sub New()
        End Sub

        Public Function Generate(ByVal Grammar As Grammar, ByVal Debug As Boolean) As String Implements ICodeGenerator.Generate
            If String.IsNullOrEmpty(Grammar.GetTemplatePath) Then
                Return Nothing
            End If
            Dim scanner As String = File.ReadAllText((Grammar.GetTemplatePath & Me.FileName))
            Dim counter As Integer = 2
            Dim tokentype As New StringBuilder
            Dim regexps As New StringBuilder
            Dim skiplist As New StringBuilder
            Dim s As TerminalSymbol
            For Each s In Grammar.SkipSymbols
                skiplist.AppendLine(("            SkipList.Add(TokenType." & s.Name & ")"))
            Next
            tokentype.AppendLine(ChrW(13) & ChrW(10) & "        'Non terminal tokens:")
            tokentype.AppendLine(Helper.Outline("_NONE_", 2, "= 0", 5))
            tokentype.AppendLine(Helper.Outline("_UNDETERMINED_", 2, "= 1", 5))
            tokentype.AppendLine(ChrW(13) & ChrW(10) & "        'Non terminal tokens:")
            Dim nts As Symbol
            For Each nts In Grammar.GetNonTerminals
                tokentype.AppendLine(Helper.Outline(nts.Name, 2, ("= " & String.Format("{0:d}", counter)), 5))
                counter += 1
            Next
            tokentype.AppendLine(ChrW(13) & ChrW(10) & "        'Terminal tokens:")
            Dim first As Boolean = True
            'Dim s As TerminalSymbol
            For Each s In Grammar.GetTerminals
                Dim vbexpr As String = s.Expression.ToString
                If vbexpr.StartsWith("@") Then
                    vbexpr = vbexpr.Substring(1)
                End If
                regexps.Append(("            regex = new Regex(" & vbexpr & ", RegexOptions.Compiled)" & ChrW(13) & ChrW(10)))
                regexps.Append(("            Patterns.Add(TokenType." & s.Name & ", regex)" & ChrW(13) & ChrW(10)))
                regexps.Append(("            Tokens.Add(TokenType." & s.Name & ")" & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10)))
                If first Then
                    first = False
                Else
                    tokentype.AppendLine("")
                End If
                tokentype.Append(Helper.Outline(s.Name, 2, ("= " & String.Format("{0:d}", counter)), 5))
                counter += 1
            Next
            scanner = scanner.Replace("<%SkipList%>", skiplist.ToString).Replace("<%RegExps%>", regexps.ToString).Replace("<%TokenType%>", tokentype.ToString)
            If Debug Then
                scanner = scanner.Replace("<%Imports%>", "Imports TinyPG.Debug").Replace("<%Namespace%>", "TinyPG.Debug").Replace("<%IToken%>", ChrW(13) & ChrW(10) & "        Implements IToken").Replace("<%ImplementsITokenStartPos%>", " Implements IToken.StartPos").Replace("<%ImplementsITokenEndPos%>", " Implements IToken.EndPos").Replace("<%ImplementsITokenLength%>", " Implements IToken.Length").Replace("<%ImplementsITokenText%>", " Implements IToken.Text").Replace("<%ImplementsITokenToString%>", " Implements IToken.ToString")
            Else
                scanner = scanner.Replace("<%Imports%>", "").Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace")).Replace("<%IToken%>", "").Replace("<%ImplementsITokenStartPos%>", "").Replace("<%ImplementsITokenEndPos%>", "").Replace("<%ImplementsITokenLength%>", "").Replace("<%ImplementsITokenText%>", "").Replace("<%ImplementsITokenToString%>", "")
            End If
            Return scanner
        End Function


        ' Properties
        Public ReadOnly Property FileName As String Implements ICodeGenerator.FileName
            Get
                Return "Scanner.vb"
            End Get
        End Property

    End Class
End Namespace

