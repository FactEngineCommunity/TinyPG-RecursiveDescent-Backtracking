Imports System
Imports System.IO
Imports System.Text
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler
Imports Microsoft.VisualBasic
Imports System.Collections.Generic

Namespace TinyPG.CodeGenerators.VBNet
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
                parser = parser.Replace("<%Imports%>", "Imports TinyPG.Debug").Replace("<%Namespace%>", "TinyPG.Debug").Replace("<%IParser%>", ChrW(13) & ChrW(10) & "        Implements IParser" & ChrW(13) & ChrW(10)).Replace("<%IParseTree%>", "IParseTree")
            Else
                parser = parser.Replace("<%Imports%>", "").Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace")).Replace("<%IParser%>", "").Replace("<%IParseTree%>", "ParseTree")
            End If
            Return parser.Replace("<%ParseNonTerminals%>", parsers.ToString)
        End Function

        Private Function GenerateParseMethod(ByVal s As NonTerminalSymbol) As String
            Dim sb As New StringBuilder
            Dim Indent As String = ParserGenerator.IndentTabs(1)

            sb.AppendLine(("        Private Function Parse" & s.Name & "(ByVal parent As ParseNode) As Boolean" & Helper.AddComment("'", ("NonTerminalSymbol: " & s.Name))))
            sb.AppendLine("            Dim tok As Token")
            sb.AppendLine("            Dim n As ParseNode")
            sb.AppendLine(String.Concat(New String() {"            Dim node As ParseNode = parent.CreateNode(m_scanner.GetToken(TokenType.", s.Name, "), """, s.Name, """)"}))
            sb.AppendLine("            Dim lbProblemSolved As Boolean = True")
            sb.AppendLine("")
            sb.AppendLine("            Dim liOriginalRange as Integer = m_scanner.StartPos")
            sb.AppendLine("            Dim liMaxRange as Integer = liOriginalRange")
            sb.AppendLine("            parent.Nodes.Add(node)")        '20210701-VM-Removed and put at end                
            sb.AppendLine("")
            sb.AppendLine("            Try")
            Dim rule As Rule
            For Each rule In s.Rules
                sb.AppendLine(Me.GenerateProductionRuleCode(s.Rules.Item(0), 3))
            Next

            'sb.AppendLine("            parent.Token.UpdateRange(node.Token)")         '20210701-VM-Removed and put at end    
            sb.AppendLine("            If m_scanner.Input.Length > (parent.Token.EndPos + 1) Then")
            'sb.AppendLine("              m_tree.Optionals.Clear()")  '20220804-VM-Commented out. Leave out.
            sb.AppendLine("            End If")
            'sb.AppendLine("            parent.Nodes.Add(node)")
            sb.AppendLine("            Finally")
            sb.AppendLine(Indent & "            If lbProblemSolved Then")
            sb.AppendLine(Indent & "                parent.Token.UpdateRange(node.Token)")
            sb.AppendLine(Indent & "                Me.MaxDistance = node.Token.EndPos")
            sb.AppendLine(Indent & "                If m_scanner.EndPos >= Me.MaxDistance Then")
            sb.AppendLine(Indent & "                   If m_tree.MaxDistance > max_tree.MaxDistance Then")
            sb.AppendLine(Indent & "                      Me.MaxDistance = m_scanner.StartPos")
            sb.AppendLine(Indent & "                      max_tree = m_tree.clone")
            sb.AppendLine(Indent & "                   End If")
            sb.AppendLine(Indent & "                End If")
            sb.AppendLine(Indent & "            Else")
            sb.AppendLine(Indent & "               m_scanner.StartPos = liOriginalRange")
            sb.AppendLine(Indent & "               parent.Nodes.Remove(node)")
            sb.AppendLine(Indent & "            End If")
            sb.AppendLine("            End Try")
            sb.AppendLine("            Return lbProblemSolved")
            sb.AppendLine(("        End Function" & Helper.AddComment("'", ("NonTerminalSymbol: " & s.Name))))
            sb.AppendLine()
            Return sb.ToString
        End Function

        Private Function GenerateProductionRuleCode(ByVal r As Rule, ByVal indentNumber As Integer, Optional abUseParent As Boolean = False) As String
            Dim i As Integer = 0
            Dim firsts As Symbols = Nothing
            Dim sb As New StringBuilder
            Dim Indent As String = ParserGenerator.IndentTabs(indentNumber)

            Select Case r.Type
                Case RuleType.Terminal
#Region "Terminal"
                    sb.AppendLine(Indent & "lbProblemSolved = True")
                    sb.AppendLine(Indent & String.Concat(New String() {Indent, "tok = m_scanner.Scan(TokenType.", r.Symbol.Name, ")", Helper.AddComment("'", ("Terminal Rule: " & r.Symbol.Name))}))
                    sb.AppendLine(Indent & "n = node.CreateNode(tok, tok.ToString() )")
                    sb.AppendLine(Indent & "node.Token.UpdateRange(tok)")
                    sb.AppendLine(Indent & "node.Nodes.Add(n)")

                    sb.AppendLine(Indent & "If m_scanner.StartPos >= Me.MaxDistance Then")
                    sb.AppendLine(Indent & "  m_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.Symbol.Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & r.Symbol.Name & Chr(34) & "))")
                    sb.AppendLine(Indent & "End If")

                    sb.AppendLine(Indent & "If tok.Type <> TokenType." & r.Symbol.Name & " Then")
                    sb.AppendLine(Indent & "  m_tree.Errors.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.Symbol.Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & r.Symbol.Name & Chr(34) & "))")
                    sb.AppendLine(Indent & "  lbProblemSolved = False")
                    sb.AppendLine(Indent & "  If liMaxRange >= Me.MaxDistance Then") '20220801-VM-Was EndPos >= 0803 Was StartPos >=
                    sb.AppendLine(Indent & "    Me.MaxDistance = m_scanner.StartPos")
                    sb.AppendLine(Indent & "    max_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.Symbol.Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & r.Symbol.Name & Chr(34) & "))")
                    sb.AppendLine(Indent & "    max_tree = m_tree.clone")
                    sb.AppendLine(Indent & "  End If")
                    sb.AppendLine(Indent & "  Return False" & ChrW(13) & ChrW(10))
                    sb.AppendLine(Indent & "Else") '20220803-VM-Added
                    sb.AppendLine(Indent & "If m_scanner.StartPos >= Me.MaxDistance Then") '20220804-VM-Keep. Provides look forward for next Optional.
                    sb.AppendLine(Indent & "    m_tree.Optionals.Clear") 'Must stay.
                    sb.AppendLine(Indent & "End If")
                    sb.AppendLine(Indent & "End If" & ChrW(13) & ChrW(10))
                    Exit Select
#End Region
                Case RuleType.NonTerminal
#Region "Non Terminal"
                    If abUseParent Then
                        sb.AppendLine(String.Concat(New String() {Indent, "Parse", r.Symbol.Name, "(parent.Nodes(parent.Nodes.Count -1))", Helper.AddComment("'", ("NonTerminal Rule: " & r.Symbol.Name))}))
                        sb.AppendLine(Indent & "If m_tree.Errors.Count > 0 Then")
                        sb.AppendLine(Indent & "  If m_scanner.EndPos > Me.MaxDistance Then") '20220730-VM-Changed to >= was >
                        sb.AppendLine(Indent & "    Me.MaxDistance = m_scanner.StartPos")
                        sb.AppendLine(Indent & "    max_tree = m_tree.clone") '20220730-VM-Commented out.
                        sb.AppendLine(Indent & "  End If")
                        '20220730-VM-Was here, commented out.
                        'sb.AppendLine(Indent & "  'If parent.Nodes(parent.Nodes.Count-1).Nodes.Count > 0 Then parent.Nodes(parent.Nodes.Count-1).Nodes.RemoveAt(Parent.Nodes(parent.Nodes.Count-1).Nodes.Count - 1)")
                        sb.AppendLine(Indent & "  lbProblemSolved = False")
                        sb.AppendLine(Indent & "Else If m_scanner.EndPos = Me.MaxDistance Then") '20220801-VM-Added
                        sb.AppendLine(Indent & "    Me.MaxDistance = m_scanner.EndPos") '20220801-VM-Added
                        sb.AppendLine(Indent & "    max_tree = m_tree.clone") '20220801-VM-Added
                        sb.AppendLine(Indent & "End If")
                    Else
                        sb.AppendLine(String.Concat(New String() {Indent, "Parse", r.Symbol.Name, "(node)", Helper.AddComment("'", ("NonTerminal Rule: " & r.Symbol.Name))}))
                        sb.AppendLine(Indent & "If m_tree.Errors.Count > 0 Then")
                        sb.AppendLine(Indent & "  If m_scanner.EndPos > Me.MaxDistance Then")
                        sb.AppendLine(Indent & "    Me.MaxDistance = m_scanner.StartPos")
                        sb.AppendLine(Indent & "    max_tree = m_tree.clone")
                        sb.AppendLine(Indent & "  End If")
                        'sb.AppendLine(Indent & "  If node.Nodes.Count > 0 Then node.Nodes.RemoveAt(node.Nodes.Count - 1)") 20220730-VM-Was here, commented out.
                        sb.AppendLine(Indent & "  lbProblemSolved = False")
                        sb.AppendLine(Indent & "End If")
                    End If

                    Exit Select
#End Region
                Case RuleType.Choice
#Region "Choice Rule"
                    i = 0
                    firsts = r.GetFirstTerminals
                    Dim s As TerminalSymbol
                    Dim rule As Rule
                    sb.Append((Indent & "tok = m_scanner.LookAhead({"))

                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine(("})" & Helper.AddComment("'", "Choice Rule")))

#Region "Optionals"
                    sb.AppendLine(Indent & "")
                    'sb.AppendLine((Indent & "    m_tree.Optionals.Clear")) '20220804-VM-Commented out.
                    For Each rule In r.Rules
                        'Dim s As TerminalSymbol
                        For Each s In rule.GetFirstTerminals
                            sb.AppendLine(Indent & "    m_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.GetFirstTerminals(0).Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & s.Name & Chr(34) & "))")
                        Next
                    Next
#End Region

                    sb.AppendLine((Indent & "Select Case tok.Type"))
#Region "Each Case"
                    sb.AppendLine((Indent & Helper.AddComment("'", "Choice Rule")))
                    For Each rule In r.Rules
                        'Dim s As TerminalSymbol
                        For Each s In rule.GetFirstTerminals
                            sb.AppendLine((Indent & "    Case TokenType." & s.Name))
                            If {RuleType.OneOrMore, RuleType.Concat, RuleType.Option, RuleType.Choice, RuleType.Terminal, RuleType.ZeroOrMore}.Contains(rule.Type) Then
                                sb.AppendLine((Indent & Me.GenerateProductionRuleCode(rule, (indentNumber + 2))))
                            Else
                                sb.AppendLine((Indent & "lbProblemSolved = " & Me.GenerateProductionRuleCode(rule, (indentNumber + 2))))
                            End If
                        Next
                    Next
#End Region
                    sb.AppendLine(Indent & "    Case Else")
#Region "Case Else"
                    sb.AppendLine(Indent & "    If m_tree.Errors.Count = 0 Then")
                    'sb.AppendLine(Indent & "    m_tree.Optionals.Clear")  '20220804-VM-Commented out.
                    sb.AppendLine(Indent & "    lbProblemSolved = False")
                    For Each rule In r.Rules
                        'Dim s As TerminalSymbol
                        For Each s In rule.GetFirstTerminals
                            sb.AppendLine((Indent & "    m_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.GetFirstTerminals(0).Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & s.Name & Chr(34) & "))"))
                        Next
                    Next
                    sb.AppendLine((Indent & "    End If"))
                    sb.AppendLine(Indent & "        m_tree.Errors.Add(new ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found."", &H0002, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos))")
                    sb.AppendLine(Indent & "        Exit Select")
#End Region
                    sb.AppendLine(Indent & "End Select" & Helper.AddComment("'", "Choice Rule"))

                    sb.AppendLine(Indent & "    If Not lbProblemSolved Then") '20210811-VM-commented out
                    sb.AppendLine(Indent & "       m_tree.Errors.Clear")  '20210811-VM-commented out
                    '
                    Dim LiInd As Integer = 0
                    For Each rule In r.Rules
#Region "Rules in Choice"
                        LiInd += 1 '20210811-VM-commented out
                        If LiInd <> 1 Then
                            sb.AppendLine(Indent & "    If Not lbProblemSolved Then")
                            sb.AppendLine(Indent & "      m_tree.Errors.Clear") '20210811-VM-commented out
                            'sb.AppendLine(Indent & "      m_scanner.StartPos = liMaxRange") '20220816-VM-Removed
                            sb.AppendLine(Indent & "      If liMaxRange > Me.MaxDistance Then")
                            sb.AppendLine(Indent & "        Me.MaxDistance = m_scanner.StartPos")
                            sb.AppendLine(Indent & "        max_tree = m_tree.clone")

                            sb.AppendLine(Indent & "      End If")
                            'sb.AppendLine(Indent & "      max_tree.Optionals.AddRange(m_tree.Optionals)") '20220803-VM-Commented out.
                            'sb.AppendLine("               node.Nodes.RemoveAt(node.Nodes.Count -1)") '20210811-VM-Commented out.

                            '20210811-VM-Commented out
                            If {RuleType.OneOrMore, RuleType.Concat, RuleType.Option, RuleType.Choice, RuleType.Terminal, RuleType.ZeroOrMore}.Contains(rule.Type) Then
                                sb.AppendLine(Indent & Me.GenerateProductionRuleCode(rule, indentNumber, True))
                            Else
                                sb.AppendLine(Indent & "lbProblemSolved = " & Me.GenerateProductionRuleCode(rule, (indentNumber + 1), True))
                            End If

                            sb.AppendLine(Indent & "    End If")
                        End If
#End Region
                    Next
                    sb.AppendLine((Indent & "    End If")) '20210811-VM-commented out

                    sb.AppendLine(Indent & " If (m_tree.Errors.Count > 0) Or Not lbProblemSolved Then") '20210811-VM-Added 'OR Not lbProblemSolved
                    'sb.AppendLine(Indent & "            parent.Token.UpdateRange(node.Token)") '20210701-VM-Removed
                    sb.AppendLine(Indent & "     parent.Nodes.Remove(node)") '20210812-VM-Removed.

                    '#Region "Optionals Replayed"
                    '                    For Each rule In r.Rules
                    '                        'Dim s As TerminalSymbol
                    '                        For Each s In rule.GetFirstTerminals
                    '                            sb.AppendLine((Indent & "    max_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.GetFirstTerminals(0).Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & s.Name & Chr(34) & "))"))
                    '                        Next
                    '                    Next
                    '#End Region

                    sb.AppendLine(Indent & "     Return False") '20210701-VM-Removed so that Choice can go to next choice
                    sb.AppendLine(Indent & " End If")

                    Exit Select
#End Region
                Case RuleType.Concat
#Region "Concat Rule"
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.AppendLine()
                        sb.AppendLine((Indent & Helper.AddComment("'", "Concat Rule")))
                        'sb.AppendLine((Indent & "lbProblemSolved = True"))                        
                        If {RuleType.OneOrMore, RuleType.Concat, RuleType.Option, RuleType.Choice, RuleType.Terminal, RuleType.ZeroOrMore}.Contains(rule.Type) Then
                            sb.AppendLine((Indent & Me.GenerateProductionRuleCode(rule, (indentNumber + 2))))
                        Else
                            sb.AppendLine((Indent & "lbProblemSolved = " & Me.GenerateProductionRuleCode(rule, (indentNumber + 2))))
                        End If
                        sb.AppendLine(Indent & "   If Not lbProblemSolved Then") '20210815-VM-Added because realistically need all the concats and was going into infinite loop.
                        sb.AppendLine(Indent & "      If m_scanner.StartPos > Me.MaxDistance Then")
                        sb.AppendLine(Indent & "        Me.MaxDistance = m_scanner.StartPos")
                        sb.AppendLine(Indent & "        max_tree = m_tree.clone")
                        sb.AppendLine(Indent & "      End If")
                        sb.AppendLine(Indent & "      m_scanner.StartPos = liMaxRange")
                        'sb.AppendLine(Indent & "      parent.Nodes.Remove(node)")
                        sb.AppendLine(Indent & "      Return False")
                        sb.AppendLine(Indent & "   Else")
                        sb.AppendLine(Indent & "      liMaxRange = m_scanner.EndPos")
                        sb.AppendLine(Indent & "      If liMaxRange > Me.MaxDistance Then")
                        sb.AppendLine(Indent & "        Me.MaxDistance = m_scanner.StartPos")
                        sb.AppendLine(Indent & "        max_tree = m_tree.clone")
                        sb.AppendLine(Indent & "      End If")
                        sb.AppendLine(Indent & "   End If")
                    Next

                    sb.AppendLine(Indent & "   If m_tree.Errors.Count > 0 Then")
                    'sb.AppendLine(Indent & "      parent.Token.UpdateRange(node.Token)") '20210701-VM-Removed                    
                    sb.AppendLine(Indent & "      Return False")
                    sb.AppendLine(Indent & "   End If")
                    Exit Select
                Case RuleType.Option
                    i = 0
                    firsts = r.GetFirstTerminals
                    sb.Append(Indent & "tok = m_scanner.LookAhead({")  '2021-VM-Changed = to is
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append("TokenType." & s.Name)
                        Else
                            sb.Append(", TokenType." & s.Name)
                        End If
                        i += 1
                    Next
                    sb.AppendLine(("})" & Helper.AddComment("'", "Option Rule")))

                    sb.AppendLine(Indent & "If m_scanner.EndPos >= Me.MaxDistance Then")
                    For Each s In r.GetFirstTerminals
                        sb.AppendLine(Indent & Indent & "    max_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & s.Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & s.Name & Chr(34) & "))")
                    Next
                    sb.AppendLine(Indent & "End If")

                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In r.GetFirstTerminals
                        If (i = 0) Then
                            sb.Append(Indent & "If tok.Type = TokenType." & s.Name)
                        Else
                            sb.Append(" Or tok.Type = TokenType." & s.Name)
                        End If
                        i += 1
                    Next
                    sb.AppendLine(" Then")
                    Dim rule As Rule
                    For Each rule In r.Rules
                        If {RuleType.OneOrMore, RuleType.Concat, RuleType.Option, RuleType.Choice, RuleType.Terminal, RuleType.ZeroOrMore}.Contains(rule.Type) Then
                            sb.AppendLine(Me.GenerateProductionRuleCode(rule, (indentNumber + 1)))
                        Else
                            sb.AppendLine("lbProblemSolved = " & Me.GenerateProductionRuleCode(rule, (indentNumber + 1)))
                        End If
                    Next
                    sb.AppendLine(Indent & "Else")
                    sb.AppendLine(Indent & Indent & "    m_tree.Optionals.Add(New ParseError(""Unexpected token '"" + tok.Text.Replace(""\n"", """") + ""' found. Expected "" + TokenType." & r.GetFirstTerminals(0).Name & ".ToString(), &H1001, 0, tok.StartPos, tok.StartPos, tok.EndPos - tok.StartPos, " & Chr(34) & r.GetFirstTerminals(0).Name & Chr(34) & "))")
                    sb.AppendLine(Indent & "End If")

                    sb.AppendLine(Indent & "If m_tree.Errors.Count > 0 Then")
                    'sb.AppendLine(Indent & "            parent.Token.UpdateRange(node.Token)") '20210701-VM-Removed
                    sb.AppendLine(Indent & "  Return False") '20210701-VM-Removed so that Choice can go to next choice
                    sb.AppendLine(Indent & "End If")
                    Exit Select
#End Region
                Case RuleType.ZeroOrMore
#Region "Zero-Or-More"
                    firsts = r.GetFirstTerminals
                    i = 0
                    sb.Append(Indent & "tok = m_scanner.LookAhead(")  '2021-VM-Changed = to is
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine(")" & Helper.AddComment("'", "ZeroOrMore Rule"))
                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(Indent & "While tok.Type = TokenType." & s.Name)
                        Else
                            sb.Append(" Or tok.Type = TokenType." & s.Name)
                        End If
                        i += 1
                    Next
                    sb.AppendLine("")
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.AppendLine(Indent & "m_tree.Errors.Clear") '20210703-VM-Added to try and fix infinite loop. See return False below.
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indentNumber + 1)))
                    Next
                    i = 0
                    sb.Append(Indent & "tok = m_scanner.LookAhead(")
                    'Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append("TokenType." & s.Name)
                        Else
                            sb.Append(", TokenType." & s.Name)
                        End If
                        i += 1
                    Next
                    sb.AppendLine(")" & Helper.AddComment("'", "ZeroOrMore Rule"))
                    sb.AppendLine(Indent & "If Not lbProblemSolved Then Exit While")
                    sb.AppendLine(Indent & "End While")

                    sb.AppendLine("            If m_tree.Errors.Count > 0 Then")
                    'sb.AppendLine(Indent & "            parent.Token.UpdateRange(node.Token)") '20210701-VM-Removed
                    sb.AppendLine(Indent & "            Return False")
                    sb.AppendLine("            End If")
                    Exit Select
#End Region
                Case RuleType.OneOrMore
#Region "One-Or-More"
                    sb.AppendLine(Indent & "Do" & Helper.AddComment("'", "OneOrMore Rule"))
                    Dim rule As Rule
                    For Each rule In r.Rules
                        sb.Append(Me.GenerateProductionRuleCode(rule, (indentNumber + 1)))
                    Next
                    i = 0
                    firsts = r.GetFirstTerminals
                    sb.Append((Indent & "    tok = m_scanner.LookAhead(")) '2021-VM-Changed = to is
                    Dim s As TerminalSymbol
                    For Each s In firsts
                        If (i = 0) Then
                            sb.Append(("TokenType." & s.Name))
                        Else
                            sb.Append((", TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine((")" & Helper.AddComment("'", "OneOrMore Rule")))
                    i = 0
                    'Dim s As TerminalSymbol
                    For Each s In r.GetFirstTerminals
                        If (i = 0) Then
                            sb.Append((Indent & "Loop While m_tree.Errors.Count = 0 And tok.Type = TokenType." & s.Name)) '20220217-VM-Put in m_tree.Errors.Count = 0
                        Else
                            sb.Append((" Or tok.Type = TokenType." & s.Name))
                        End If
                        i += 1
                    Next
                    sb.AppendLine(If(Helper.AddComment("'", "OneOrMore Rule") <> Nothing, Helper.AddComment("'", "OneOrMore Rule"), ""))

                    sb.AppendLine("            If m_tree.Errors.Count > 0 Then")
                    'sb.AppendLine(Indent & "            parent.Token.UpdateRange(node.Token)") '20210701-VM-Removed
                    sb.AppendLine(Indent & "            Return False") '20210701-VM-Removed so that Choice can go to next choice
                    sb.AppendLine("            End If")
                    Exit Select
#End Region
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
                Return "Parser.vb"
            End Get
        End Property

    End Class
End Namespace

