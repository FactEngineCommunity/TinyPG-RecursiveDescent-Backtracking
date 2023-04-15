Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports TinyPG

Namespace TinyPG.Compiler
    Public Class GrammarNode
        Inherits ParseTree
        ' Methods
        Protected Sub New()
        End Sub

        Protected Sub New(ByVal token As Token, ByVal [text] As String)
            MyBase.Token = token
            MyBase.text = [text]
            MyBase.nodesField = New List(Of ParseNode)
        End Sub

        Public Overrides Function CreateNode(ByVal token As Token, ByVal [text] As String) As ParseNode
            Return New GrammarNode(token, [text]) With { _
                .Parent = Me _
            }
        End Function

        Protected Overrides Function EvalAttribute(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim grammar As Grammar = DirectCast(paramlist(0), Grammar)
            Dim symbol As Symbol = DirectCast(paramlist(1), Symbol)
            Dim node As GrammarNode = DirectCast(paramlist(2), GrammarNode)
            If symbol.Attributes.ContainsKey(node.Nodes.Item(1).Token.Text) Then
                tree.Errors.Add(New ParseError(("Attribute already defined for this symbol: " & node.Nodes.Item(1).Token.Text), &H1039, node.Nodes.Item(1)))
                Return Nothing
            End If
            symbol.Attributes.Add(node.Nodes.Item(1).Token.Text, DirectCast(Me.EvalParams(tree, New Object() { node }), Object()))
            Select Case node.Nodes.Item(1).Token.Text
                Case "Skip"
                    If TypeOf symbol Is TerminalSymbol Then
                        grammar.SkipSymbols.Add(symbol)
                    Else
                        tree.Errors.Add(New ParseError(("Attribute for Non terminal rule not allowed: " & node.Nodes.Item(1).Token.Text), &H1035, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length))
                    End If
                    Return symbol
                Case "Color"
                    If TypeOf symbol Is NonTerminalSymbol Then
                        tree.Errors.Add(New ParseError(("Attribute for Non terminal rule not allowed: " & node.Nodes.Item(1).Token.Text), &H1035, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length))
                    End If
                    If ((symbol.Attributes.Item("Color").Length <> 1) AndAlso (symbol.Attributes.Item("Color").Length <> 3)) Then
                        tree.Errors.Add(New ParseError(("Attribute " & node.Nodes.Item(1).Token.Text & " has too many or missing parameters"), &H103A, node.Nodes.Item(1)))
                    End If
                    Dim i As Integer
                    For i = 0 To symbol.Attributes.Item("Color").Length - 1
                        If TypeOf symbol.Attributes.Item("Color")(i) Is String Then
                            tree.Errors.Add(New ParseError(("Parameter " & node.Nodes.Item(3).Nodes.Item((i * 2)).Nodes.Item(0).Token.Text & " is of incorrect type"), &H103A, node.Nodes.Item(3).Nodes.Item((i * 2)).Nodes.Item(0)))
                            Return symbol
                        End If
                    Next i
                    Return symbol
            End Select
            tree.Errors.Add(New ParseError(("Attribute not supported: " & node.Nodes.Item(1).Token.Text), &H1036, node.Nodes.Item(1)))
            Return symbol
        End Function

        Protected Overrides Function EvalConcatRule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            If (MyBase.Nodes.Count = 1) Then
                Return MyBase.Nodes.Item(0).Eval(tree, paramlist)
            End If
            Dim concatRule As New Rule(RuleType.Concat)
            Dim i As Integer
            For i = 0 To MyBase.Nodes.Count - 1
                Dim rule As Rule = DirectCast(MyBase.Nodes.Item(i).Eval(tree, paramlist), Rule)
                concatRule.Rules.Add(rule)
            Next i
            Return concatRule
        End Function

        Protected Overrides Function EvalDirective(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim g As Grammar = DirectCast(paramlist(0), Grammar)
            Dim node As GrammarNode = DirectCast(paramlist(1), GrammarNode)
            Dim name As String = node.Nodes.Item(1).Token.Text
            'Dim CS$4$0001 As String = name
            If ((Not name Is Nothing) AndAlso (((name = "TinyPG") OrElse (name = "Parser")) OrElse (((name = "Scanner") OrElse (name = "ParseTree")) OrElse (name = "TextHighlighter")))) Then
                If (Not g.Directives.Find(name) Is Nothing) Then
                    tree.Errors.Add(New ParseError(("Directive '" & name & "' is already defined"), &H1030, node.Nodes.Item(1)))
                    Return Nothing
                End If
            Else
                tree.Errors.Add(New ParseError(("Directive '" & name & "' is not supported"), &H1031, node.Nodes.Item(1)))
            End If
            Dim directive As New Directive(name)
            g.Directives.Add(directive)
            Dim n As ParseNode
            For Each n In node.Nodes
                If (n.Token.Type = TokenType.NameValue) Then
                    Me.EvalNameValue(tree, New Object() { g, directive, n })
                End If
            Next
            Return Nothing
        End Function

        Protected Overrides Function EvalExtProduction(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Return MyBase.Nodes.Item((MyBase.Nodes.Count - 1)).Eval(tree, paramlist)
        End Function

        Protected Overrides Function EvalNameValue(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim grammer As Grammar = DirectCast(paramlist(0), Grammar)
            Dim directive As Directive = DirectCast(paramlist(1), Directive)
            Dim node As GrammarNode = DirectCast(paramlist(2), GrammarNode)
            Dim key As String = node.Nodes.Item(0).Token.Text
            Dim value As String = node.Nodes.Item(2).Token.Text.Substring(1, (node.Nodes.Item(2).Token.Text.Length - 2))
            If value.StartsWith("""") Then
                value = value.Substring(1)
            End If
            directive.Item(key) = value
            Dim names As New List(Of String)(New String() { "Namespace", "OutputPath", "TemplatePath" })
            Dim languages As New List(Of String)(New String() { "c#", "cs", "csharp", "vb", "vb.net", "vbnet", "visualbasic" })
            Dim directiveName As String = directive.Name
            If (Not directiveName Is Nothing) Then
                If Not (directiveName = "TinyPG") Then
                    If (((directiveName = "Parser") OrElse (directiveName = "Scanner")) OrElse ((directiveName = "ParseTree") OrElse (directiveName = "TextHighlighter"))) Then
                        names.Add("Generate")
                        GoTo Label_02D5
                    End If
                Else
                    names.Add("Namespace")
                    names.Add("OutputPath")
                    names.Add("TemplatePath")
                    names.Add("Language")
                    If ((key = "TemplatePath") AndAlso (grammer.GetTemplatePath Is Nothing)) Then
                        tree.Errors.Add(New ParseError(("Template path '" & value & "' does not exist"), &H1060, node.Nodes.Item(2)))
                    End If
                    If ((key = "OutputPath") AndAlso (grammer.GetOutputPath Is Nothing)) Then
                        tree.Errors.Add(New ParseError(("Output path '" & value & "' does not exist"), &H1061, node.Nodes.Item(2)))
                    End If
                    If ((key = "Language") AndAlso Not languages.Contains(value.ToLower(CultureInfo.InvariantCulture))) Then
                        tree.Errors.Add(New ParseError(("Language '" & value & "' is not supported"), &H1062, node.Nodes.Item(2)))
                    End If
                    GoTo Label_02D5
                End If
            End If
            Return Nothing
        Label_02D5:
            If Not names.Contains(key) Then
                tree.Errors.Add(New ParseError(("Directive attribute '" & key & "' is not supported"), &H1034, node.Nodes.Item(0)))
            End If
            Return Nothing
        End Function

        Protected Overrides Function EvalParam(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim node As GrammarNode = DirectCast(paramlist(0), GrammarNode)
            Try 
                Select Case node.Nodes.Item(0).Token.Type
                    Case TokenType.INTEGER
                        Return Convert.ToInt32(node.Nodes.Item(0).Token.Text)
                    Case TokenType.HEX
                        Return Long.Parse(node.Nodes.Item(0).Token.Text.Substring(2), NumberStyles.HexNumber)
                    Case TokenType.STRING
                        Return node.Nodes.Item(0).Token.Text
                End Select
                tree.Errors.Add(New ParseError(("Attribute parameter is not a valid value: " & node.Token.Text), &H1037, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length))
            Catch exception1 As Exception
                tree.Errors.Add(New ParseError(("Attribute parameter is not a valid value: " & node.Token.Text), &H1038, node))
            End Try
            Return Nothing
        End Function

        Protected Overrides Function EvalParams(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim node As GrammarNode = DirectCast(paramlist(0), GrammarNode)
            If (node.Nodes.Count < 4) Then
                Return Nothing
            End If
            If (node.Nodes.Item(3).Token.Type <> TokenType.Params) Then
                Return Nothing
            End If
            Dim parms As GrammarNode = DirectCast(node.Nodes.Item(3), GrammarNode)
            Dim objects As New List(Of Object)
            Dim i As Integer = 0
            Do While (i < parms.Nodes.Count)
                objects.Add(Me.EvalParam(tree, New Object() { parms.Nodes.Item(i) }))
                i = (i + 2)
            Loop
            Return objects.ToArray
        End Function

        Protected Overrides Function EvalProduction(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim g As Grammar = DirectCast(paramlist(0), Grammar)
            If (MyBase.Nodes.Item(2).Nodes.Item(0).Token.Type = TokenType.STRING) Then
                Dim term As TerminalSymbol = TryCast(g.Symbols.Find(MyBase.Nodes.Item(0).Token.Text),TerminalSymbol)
                If (term Is Nothing) Then
                    tree.Errors.Add(New ParseError(("Symbol '" & MyBase.Nodes.Item(0).Token.Text & "' is not declared. "), &H1040, MyBase.Nodes.Item(0)))
                End If
                Return g
            End If
            Dim nts As NonTerminalSymbol = TryCast(g.Symbols.Find(MyBase.Nodes.Item(0).Token.Text),NonTerminalSymbol)
            If (nts Is Nothing) Then
                tree.Errors.Add(New ParseError(("Symbol '" & MyBase.Nodes.Item(0).Token.Text & "' is not declared. "), &H1041, MyBase.Nodes.Item(0)))
            End If
            Dim r As Rule = DirectCast(MyBase.Nodes.Item(2).Eval(tree, New Object() { g, nts }), Rule)
            If (Not nts Is Nothing) Then
                nts.Rules.Add(r)
            End If
            If (MyBase.Nodes.Item(3).Token.Type = TokenType.CODEBLOCK) Then
                Dim codeblock As String = MyBase.Nodes.Item(3).Token.Text
                nts.CodeBlock = codeblock
                Me.ValidateCodeBlock(tree, nts, MyBase.Nodes.Item(3))
                codeblock = codeblock.Substring(1, (codeblock.Length - 3)).Trim
                nts.CodeBlock = codeblock
            End If
            Return g
        End Function

        Protected Overrides Function EvalRule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Return MyBase.Nodes.Item(0).Eval(tree, paramlist)
        End Function

        Protected Overrides Function EvalStart(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim terminal As TerminalSymbol = Nothing
            Dim StartFound As Boolean = False
            Dim g As New Grammar
            Dim n As ParseNode
            For Each n In MyBase.Nodes
                If (n.Token.Type = TokenType.Directive) Then
                    Me.EvalDirective(tree, New Object() { g, n })
                End If
                If (n.Token.Type = TokenType.ExtProduction) Then
                    Dim i As Integer
                    If (n.Nodes.Item((n.Nodes.Count - 1)).Nodes.Item(2).Nodes.Item(0).Token.Type = TokenType.STRING) Then
                        Try 
                            terminal = New TerminalSymbol(n.Nodes.Item((n.Nodes.Count - 1)).Nodes.Item(0).Token.Text, n.Nodes.Item((n.Nodes.Count - 1)).Nodes.Item(2).Nodes.Item(0).Token.Text)
                            i = 0
                            Do While (i < (n.Nodes.Count - 1))
                                If (n.Nodes.Item(i).Token.Type = TokenType.Attribute) Then
                                    Me.EvalAttribute(tree, New Object() { g, terminal, n.Nodes.Item(i) })
                                End If
                                i += 1
                            Loop
                        Catch ex As Exception
                            tree.Errors.Add(New ParseError(("regular expression for '" & n.Nodes.Item((n.Nodes.Count - 1)).Nodes.Item(0).Token.Text & "' results in error: " & ex.Message), &H1020, 0, n.Nodes.Item(0).Token.StartPos, n.Nodes.Item(0).Token.StartPos, (n.Nodes.Item(0).Token.EndPos - n.Nodes.Item(0).Token.StartPos)))
                            Continue For
                        End Try
                        If (terminal.Name = "Start") Then
                            tree.Errors.Add(New ParseError("'Start' symbol cannot be a regular expression.", &H1021, 0, n.Nodes.Item(0).Token.StartPos, n.Nodes.Item(0).Token.StartPos, (n.Nodes.Item(0).Token.EndPos - n.Nodes.Item(0).Token.StartPos)))
                        End If
                        If (g.Symbols.Find(terminal.Name) Is Nothing) Then
                            g.Symbols.Add(terminal)
                        Else
                            tree.Errors.Add(New ParseError(("Terminal already declared: " & terminal.Name), &H1022, 0, n.Nodes.Item(0).Token.StartPos, n.Nodes.Item(0).Token.StartPos, (n.Nodes.Item(0).Token.EndPos - n.Nodes.Item(0).Token.StartPos)))
                        End If
                    Else
                        Dim nts As New NonTerminalSymbol(n.Nodes.Item((n.Nodes.Count - 1)).Nodes.Item(0).Token.Text)
                        If (g.Symbols.Find(nts.Name) Is Nothing) Then
                            g.Symbols.Add(nts)
                        Else
                            tree.Errors.Add(New ParseError(("Non terminal already declared: " & nts.Name), &H1023, 0, n.Nodes.Item(0).Token.StartPos, n.Nodes.Item(0).Token.StartPos, (n.Nodes.Item(0).Token.EndPos - n.Nodes.Item(0).Token.StartPos)))
                        End If

                        For i = 0 To (n.Nodes.Count - 1) - 1
                            If (n.Nodes.Item(i).Token.Type = TokenType.Attribute) Then
                                Me.EvalAttribute(tree, New Object() { g, nts, n.Nodes.Item(i) })
                            End If
                        Next i
                        If (nts.Name = "Start") Then
                            StartFound = True
                        End If
                    End If
                End If
            Next
            If Not StartFound Then
                tree.Errors.Add(New ParseError("The grammar requires 'Start' to be a production rule.", &H24, 0, 0, 0, 0))
                Return g
            End If

            For Each n In MyBase.Nodes
                If (n.Token.Type = TokenType.ExtProduction) Then
                    n.Eval(tree, New Object() { g })
                End If
            Next
            Return g
        End Function

        Protected Overrides Function EvalSubrule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            If (MyBase.Nodes.Count = 1) Then
                Return MyBase.Nodes.Item(0).Eval(tree, paramlist)
            End If
            Dim choiceRule As New Rule(RuleType.Choice)
            Dim i As Integer = 0
            Do While (i < MyBase.Nodes.Count)
                Dim rule As Rule = DirectCast(MyBase.Nodes.Item(i).Eval(tree, paramlist), Rule)
                choiceRule.Rules.Add(rule)
                i = (i + 2)
            Loop
            Return choiceRule
        End Function

        Protected Overrides Function EvalSymbol(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Dim g As Grammar
            Dim s As Symbol
            Dim last As ParseNode = MyBase.Nodes.Item((MyBase.Nodes.Count - 1))
            If (last.Token.Type = TokenType.UNARYOPER) Then
                Dim unaryRule As Rule
                Dim oper As String = last.Token.Text.Trim
                If (oper = "*") Then
                    unaryRule = New Rule(RuleType.ZeroOrMore)
                ElseIf (oper = "+") Then
                    unaryRule = New Rule(RuleType.OneOrMore)
                Else
                    If (oper <> "?") Then
                        Throw New NotImplementedException("unknown unary operator")
                    End If
                    unaryRule = New Rule(RuleType.Option)
                End If
                If (MyBase.Nodes.Item(0).Token.Type = TokenType.BRACKETOPEN) Then
                    Dim rule As Rule = DirectCast(MyBase.Nodes.Item(1).Eval(tree, paramlist), Rule)
                    unaryRule.Rules.Add(rule)
                    Return unaryRule
                End If
                g = DirectCast(paramlist(0), Grammar)
                If (MyBase.Nodes.Item(0).Token.Type = TokenType.IDENTIFIER) Then
                    s = g.Symbols.Find(MyBase.Nodes.Item(0).Token.Text)
                    If (s Is Nothing) Then
                        tree.Errors.Add(New ParseError(("Symbol '" & MyBase.Nodes.Item(0).Token.Text & "' is not declared. "), &H1042, MyBase.Nodes.Item(0)))
                    End If
                    Dim r As New Rule(s)
                    unaryRule.Rules.Add(r)
                End If
                Return unaryRule
            End If
            If (MyBase.Nodes.Item(0).Token.Type = TokenType.BRACKETOPEN) Then
                Return MyBase.Nodes.Item(1).Eval(tree, paramlist)
            End If
            g = DirectCast(paramlist(0), Grammar)
            s = g.Symbols.Find(MyBase.Nodes.Item(0).Token.Text)
            If (s Is Nothing) Then
                tree.Errors.Add(New ParseError(("Symbol '" & MyBase.Nodes.Item(0).Token.Text & "' is not declared."), &H1043, MyBase.Nodes.Item(0)))
            End If
            Return New Rule(s)
        End Function

        Private Sub ValidateCodeBlock(ByVal tree As ParseTree, ByVal nts As NonTerminalSymbol, ByVal node As ParseNode)
            If (Not nts Is Nothing) Then
                Dim codeblock As String = nts.CodeBlock
                Dim var As New Regex("\$(?<var>[a-zA-Z_0-9]+)(\[(?<index>[^]]+)\])?", RegexOptions.Compiled)
                Dim symbols As Symbols = nts.DetermineProductionSymbols
                Dim matches As MatchCollection = var.Matches(codeblock)
                Dim match As Match
                For Each match In matches
                    If (symbols.Find(match.Groups.Item("var").Value) Is Nothing) Then
                        tree.Errors.Add(New ParseError(("Variable $" & match.Groups.Item("var").Value & " cannot be matched."), &H1016, 0, (node.Token.StartPos + match.Groups.Item("var").Index), (node.Token.StartPos + match.Groups.Item("var").Index), match.Groups.Item("var").Length))
                        Exit For
                    End If
                Next
            End If
        End Sub

    End Class
End Namespace

