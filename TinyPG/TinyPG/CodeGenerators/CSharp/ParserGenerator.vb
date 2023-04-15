Imports System
Imports System.IO
Imports System.Text
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler
Imports Microsoft.VisualBasic

Namespace TinyPG.CodeGenerators.CSharp
    Public Class ParserGenerator
        Implements ICodeGenerator
        ' Methods
        Friend Sub New()
        End Sub

        Public Function Generate(ByVal Grammar As Grammar, ByVal Debug As Boolean) As String Implements ICodeGenerator.Generate
            If String.IsNullOrEmpty(Grammar.GetTemplatePath) Then
                Return Nothing
            End If
            Dim parsers As New StringBuilder
            Dim parser As String = File.ReadAllText((Grammar.GetTemplatePath & Me.FileName))
            Dim s As NonTerminalSymbol
            For Each s In Grammar.GetNonTerminals
                Dim method As String = Me.GenerateParseMethod(s)
                parsers.Append(method)
            Next
            If Debug Then
                parser = parser.Replace("<%Namespace%>", "TinyPG.Debug").Replace("<%IParser%>", " : TinyPG.Debug.IParser").Replace("<%IParseTree%>", "TinyPG.Debug.IParseTree")
            Else
                parser = parser.Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace")).Replace("<%IParser%>", "").Replace("<%IParseTree%>", "ParseTree")
            End If
            Return parser.Replace("<%ParseNonTerminals%>", parsers.ToString)
        End Function

        Private Function GenerateParseMethod(ByVal symbol As NonTerminalSymbol) As String
            Dim sb As New StringBuilder
            sb.AppendLine(("        private void Parse" & symbol.Name & "(ParseNode parent)" & Helper.AddComment(("NonTerminalSymbol: " & symbol.Name))))
            sb.AppendLine("        {")
            sb.AppendLine("            Token tok;")
            sb.AppendLine("            ParseNode n;")
            sb.AppendLine(String.Concat(New String() {"            ParseNode node = parent.CreateNode(scanner.GetToken(TokenType.", symbol.Name, "), """, symbol.Name, """);"}))
            sb.AppendLine("            parent.Nodes.Add(node);")
            sb.AppendLine("")
            Dim rule As Rule
            For Each rule In symbol.Rules
                sb.AppendLine(Me.GenerateProductionRuleCode(symbol.Rules.Item(0), 3))
            Next
            sb.AppendLine("            parent.Token.UpdateRange(node.Token);")
            sb.AppendLine(("        }" & Helper.AddComment(("NonTerminalSymbol: " & symbol.Name))))
            sb.AppendLine()
            Return sb.ToString
        End Function

        Private Function GenerateProductionRuleCode(ByVal r As Rule, ByVal indent As Integer) As String
            Dim i As Integer = 0
            Dim firsts As Symbols = Nothing
            Dim sb As New StringBuilder
            Dim code As String = ParserGenerator.IndentTabs(indent)
            Select Case r.Type
                Case RuleType.Terminal
                    sb.AppendLine(String.Concat(New String() {indent.ToString, "tok = scanner.Scan(TokenType.", r.Symbol.Name, ");", Helper.AddComment(("Terminal Rule: " & r.Symbol.Name))}))
                    sb.AppendLine((code & "n = node.CreateNode(tok, tok.ToString() );"))
                    sb.AppendLine((code & "node.Token.UpdateRange(tok);"))
                    sb.AppendLine((code & "node.Nodes.Add(n);"))
                    sb.AppendLine((code & "if (tok.Type != TokenType." & r.Symbol.Name & ") {"))
                    sb.AppendLine((code & "    tree.Errors.Add(new ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.Symbol.Name & ".ToString(), 0x1001, 0, tok.StartPos, tok.StartPos, tok.Length));"))
                    sb.AppendLine((code & "    return;"))
                    sb.AppendLine((code & "}"))
                    Exit Select
                Case RuleType.NonTerminal
                    sb.AppendLine(String.Concat(New String() {indent.ToString, "Parse", r.Symbol.Name, "(node);", Helper.AddComment(("NonTerminal Rule: " & r.Symbol.Name))}))
                    Exit Select
                Case RuleType.Choice
                    i = 0
                    firsts = r.GetFirstTerminals
                    sb.Append((code & "tok = scanner.LookAhead("))
                    Dim symbol As TerminalSymbol
                    For Each symbol In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & symbol.Name))
                        Else
                            sb.Append((", TokenType." & symbol.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("Choice Rule")))
                    sb.AppendLine((code & "switch (tok.Type)"))
                    sb.AppendLine((code & "{" & Helper.AddComment("Choice Rule")))
                    Dim rule As Rule
                    For Each rule In r.Rules
                        'Dim symbol1 As TerminalSymbol
                        For Each symbol In rule.GetFirstTerminals
                            sb.AppendLine((code & "    case TokenType." & symbol.Name & ":"))
                        Next
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indent + 2)))
                        sb.AppendLine((code & "        break;"))
                    Next
                    sb.AppendLine((code & "    default:"))
                    sb.AppendLine((code & "        tree.Errors.Add(new ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found."", 0x0002, 0, tok.StartPos, tok.StartPos, tok.Length));"))
                    sb.AppendLine((code & "        break;"))
                    sb.AppendLine((code & "}" & Helper.AddComment("Choice Rule")))
                    Exit Select
                Case RuleType.Concat
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.AppendLine()
                        sb.AppendLine((code & Helper.AddComment("Concat Rule")))
                        sb.Append(Me.GenerateProductionRuleCode(rule, indent))
                    Next
                    Exit Select
                Case RuleType.Option
                    i = 0
                    firsts = r.GetFirstTerminals
                    sb.Append((code & "tok = scanner.LookAhead("))
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("Option Rule")))
                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append((code & "if (tok.Type == TokenType." & s.Name))
                        Else
                            sb.Append((ChrW(13) & ChrW(10) & code & "    || tok.Type == TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine(")")
                    sb.AppendLine((code & "{"))
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indent + 1)))
                    Next
                    sb.AppendLine((code & "}"))
                    Exit Select
                Case RuleType.ZeroOrMore
                    firsts = r.GetFirstTerminals
                    i = 0
                    sb.Append((code & "tok = scanner.LookAhead("))
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("ZeroOrMore Rule")))
                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append((code & "while (tok.Type == TokenType." & s.Name))
                        Else
                            sb.Append((ChrW(13) & ChrW(10) & code & "    || tok.Type == TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine(")")
                    sb.AppendLine((code & "{"))
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indent + 1)))
                    Next
                    i = 0
                    sb.Append((code & "tok = scanner.LookAhead("))
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("ZeroOrMore Rule")))
                    sb.AppendLine((code & "}"))
                    Exit Select
                Case RuleType.OneOrMore
                    sb.AppendLine((code & "do {" & Helper.AddComment("OneOrMore Rule")))
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indent + 1)))
                    Next
                    i = 0
                    firsts = r.GetFirstTerminals
                    sb.Append((code & "    tok = scanner.LookAhead("))
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("OneOrMore Rule")))
                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append((code & "} while (tok.Type == TokenType." & s.Name))
                        Else
                            sb.Append((ChrW(13) & ChrW(10) & code & "    || tok.Type == TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((");" & Helper.AddComment("OneOrMore Rule")))
                    Exit Select
            End Select
            Return sb.ToString
        End Function

        Public Shared Function IndentTabs(ByVal indent As Integer) As String
            Dim t As String = ""
            Dim i As Integer
            For i = 0 To indent - 1
                t = (t & "    ")
            Next i
            Return t
        End Function


        ' Properties
        Public ReadOnly Property FileName As String Implements ICodeGenerator.FileName
            Get
                Return "Parser.cs"
            End Get
        End Property

    End Class
End Namespace

