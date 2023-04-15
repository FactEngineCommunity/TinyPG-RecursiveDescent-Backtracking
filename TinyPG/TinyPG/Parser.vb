Imports System
Imports Microsoft.VisualBasic

Namespace TinyPG
    Public Class Parser
        ' Methods
        Public Sub New(ByVal scanner As Scanner)
            Me.scanner = scanner
        End Sub

        Public Function Parse(ByVal input As String) As ParseTree
            Me.tree = New ParseTree
            Return Me.Parse(input, Me.tree)
        End Function

        Public Function Parse(ByVal input As String, ByVal tree As ParseTree) As ParseTree
            Me.scanner.Init(input)
            Me.tree = tree
            Me.ParseStart(tree)
            tree.Skipped = Me.scanner.Skipped
            Return tree
        End Function

        Private Sub ParseAttribute(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Attribute), "Attribute")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() { TokenType.SQUAREOPEN })
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.Token.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.SQUAREOPEN) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.SQUAREOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.Scan(New TokenType() {TokenType.IDENTIFIER})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.IDENTIFIER) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.IDENTIFIER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                Else
                    If (Me.scanner.LookAhead(New TokenType() {TokenType.BRACKETOPEN}).Type = TokenType.BRACKETOPEN) Then
                        tok = Me.scanner.Scan(New TokenType() {TokenType.BRACKETOPEN})
                        n = node.CreateNode(tok, tok.ToString)
                        node.Token.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type <> TokenType.BRACKETOPEN) Then
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.BRACKETOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        End If
                        tok = Me.scanner.LookAhead(New TokenType() {TokenType.INTEGER, TokenType.DOUBLE, TokenType.STRING, TokenType.HEX})
                        If ((((tok.Type = TokenType.INTEGER) OrElse (tok.Type = TokenType.DOUBLE)) OrElse (tok.Type = TokenType.STRING)) OrElse (tok.Type = TokenType.HEX)) Then
                            Me.ParseParams(node)
                        End If
                        tok = Me.scanner.Scan(New TokenType() {TokenType.BRACKETCLOSE})
                        n = node.CreateNode(tok, tok.ToString)
                        node.Token.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type <> TokenType.BRACKETCLOSE) Then
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.BRACKETCLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        End If
                    End If
                    tok = Me.scanner.Scan(New TokenType() {TokenType.SQUARECLOSE})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.SQUARECLOSE) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.SQUARECLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Else
                        parent.Token.UpdateRange(node.Token)
                    End If
                End If
            End If
        End Sub

        Private Sub ParseConcatRule(ByVal parent As ParseNode)
            Dim tok As Token
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.ConcatRule), "ConcatRule")
            parent.Nodes.Add(node)
            Do
                Me.ParseSymbol(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.IDENTIFIER, TokenType.BRACKETOPEN})
            Loop While ((tok.Type = TokenType.IDENTIFIER) OrElse (tok.Type = TokenType.BRACKETOPEN))
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseDirective(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Directive), "Directive")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() {TokenType.DIRECTIVEOPEN})
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.Token.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.DIRECTIVEOPEN) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVEOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.Scan(New TokenType() {TokenType.IDENTIFIER})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.IDENTIFIER) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.IDENTIFIER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                Else
                    tok = Me.scanner.LookAhead(New TokenType() {TokenType.IDENTIFIER})
                    Do While (tok.Type = TokenType.IDENTIFIER)
                        Me.ParseNameValue(node)
                        tok = Me.scanner.LookAhead(New TokenType() {TokenType.IDENTIFIER})
                    Loop
                    tok = Me.scanner.Scan(New TokenType() {TokenType.DIRECTIVECLOSE})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.DIRECTIVECLOSE) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVECLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Else
                        parent.Token.UpdateRange(node.Token)
                    End If
                End If
            End If
        End Sub

        Private Sub ParseExtProduction(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.ExtProduction), "ExtProduction")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.SQUAREOPEN})
            Do While (tok.Type = TokenType.SQUAREOPEN)
                Me.ParseAttribute(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.SQUAREOPEN})
            Loop
            Me.ParseProduction(node)
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseNameValue(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.NameValue), "NameValue")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() {TokenType.IDENTIFIER})
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.Token.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.IDENTIFIER) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.IDENTIFIER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.Scan(New TokenType() {TokenType.ASSIGN})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.ASSIGN) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ASSIGN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                Else
                    tok = Me.scanner.Scan(New TokenType() {TokenType.STRING})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.STRING) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.STRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Else
                        parent.Token.UpdateRange(node.Token)
                    End If
                End If
            End If
        End Sub

        Private Sub ParseParam(ByVal parent As ParseNode)
            Dim n As ParseNode
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Param), "Param")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.INTEGER, TokenType.DOUBLE, TokenType.STRING, TokenType.HEX})
            Select Case tok.Type
                Case TokenType.INTEGER
                    tok = Me.scanner.Scan(New TokenType() {TokenType.INTEGER})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.INTEGER) Then
                        Exit Select
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.INTEGER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                Case TokenType.DOUBLE
                    tok = Me.scanner.Scan(New TokenType() {TokenType.DOUBLE})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.DOUBLE) Then
                        Exit Select
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOUBLE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                Case TokenType.HEX
                    tok = Me.scanner.Scan(New TokenType() {TokenType.HEX})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.HEX) Then
                        Exit Select
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.HEX.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                Case TokenType.STRING
                    tok = Me.scanner.Scan(New TokenType() {TokenType.STRING})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.STRING) Then
                        Exit Select
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.STRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                Case Else
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Exit Select
            End Select
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseParams(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Params), "Params")
            parent.Nodes.Add(node)
            Me.ParseParam(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.COMMA})
            Do While (tok.Type = TokenType.COMMA)
                tok = Me.scanner.Scan(New TokenType() {TokenType.COMMA})
                Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.COMMA) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.COMMA.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                End If
                Me.ParseParam(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.COMMA})
            Loop
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseProduction(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Production), "Production")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() {TokenType.IDENTIFIER})
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.Token.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.IDENTIFIER) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.IDENTIFIER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.Scan(New TokenType() {TokenType.ARROW})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.ARROW) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ARROW.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                Else
                    Me.ParseRule(node)
                    tok = Me.scanner.LookAhead(New TokenType() {TokenType.CODEBLOCK, TokenType.SEMICOLON})
                    Dim targetToken As TokenType = tok.Type
                    If (targetToken <> TokenType.CODEBLOCK) Then
                        If (targetToken = TokenType.SEMICOLON) Then
                            tok = Me.scanner.Scan(New TokenType() {TokenType.SEMICOLON})
                            n = node.CreateNode(tok, tok.ToString)
                            node.Token.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type <> TokenType.SEMICOLON) Then
                                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.SEMICOLON.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                                Return
                            End If
                        Else
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                        End If
                        GoTo Label_0391
                    End If
                    tok = Me.scanner.Scan(New TokenType() {TokenType.CODEBLOCK})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.CODEBLOCK) Then
                        GoTo Label_0391
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.CODEBLOCK.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                End If
            End If
            Return
Label_0391:
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseRule(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Rule), "Rule")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.STRING, TokenType.IDENTIFIER, TokenType.BRACKETOPEN})
            Select Case tok.Type
                Case TokenType.BRACKETOPEN, TokenType.IDENTIFIER
                    Me.ParseSubrule(node)
                    Exit Select
                Case TokenType.STRING
                    tok = Me.scanner.Scan(New TokenType() {TokenType.STRING})
                    Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.STRING) Then
                        Exit Select
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.STRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                Case Else
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Exit Select
            End Select
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseStart(ByVal parent As ParseNode)
            Dim tok As Token
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Start), "Start")
            parent.Nodes.Add(node)
            tok = Me.scanner.LookAhead(New TokenType() {TokenType.DIRECTIVEOPEN})
            Do While (tok.Type = TokenType.DIRECTIVEOPEN)
                Me.ParseDirective(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.DIRECTIVEOPEN})
            Loop
            tok = Me.scanner.LookAhead(New TokenType() {TokenType.SQUAREOPEN, TokenType.IDENTIFIER})
            Do While ((tok.Type = TokenType.SQUAREOPEN) OrElse (tok.Type = TokenType.IDENTIFIER))
                Me.ParseExtProduction(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.SQUAREOPEN, TokenType.IDENTIFIER})
            Loop
            tok = Me.scanner.Scan(New TokenType() {TokenType.EOF})
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.Token.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.EOF) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.EOF.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                parent.Token.UpdateRange(node.Token)
            End If
        End Sub

        Private Sub ParseSubrule(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Subrule), "Subrule")
            parent.Nodes.Add(node)
            Me.ParseConcatRule(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.PIPE})
            Do While (tok.Type = TokenType.PIPE)
                tok = Me.scanner.Scan(New TokenType() {TokenType.PIPE})
                Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.PIPE) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.PIPE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                End If
                Me.ParseConcatRule(node)
                tok = Me.scanner.LookAhead(New TokenType() {TokenType.PIPE})
            Loop
            parent.Token.UpdateRange(node.Token)
        End Sub

        Private Sub ParseSymbol(ByVal parent As ParseNode)
            Dim n As ParseNode
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Symbol), "Symbol")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.LookAhead(New TokenType() {TokenType.IDENTIFIER, TokenType.BRACKETOPEN})
            Dim targetToken As TokenType = tok.Type
            If (targetToken <> TokenType.BRACKETOPEN) Then
                If (targetToken <> TokenType.IDENTIFIER) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                    GoTo Label_02DF
                End If
                tok = Me.scanner.Scan(New TokenType() {TokenType.IDENTIFIER})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type = TokenType.IDENTIFIER) Then
                    GoTo Label_02DF
                End If
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.IDENTIFIER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.Scan(New TokenType() {TokenType.BRACKETOPEN})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.BRACKETOPEN) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.BRACKETOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                Else
                    Me.ParseSubrule(node)
                    tok = Me.scanner.Scan(New TokenType() {TokenType.BRACKETCLOSE})
                    n = node.CreateNode(tok, tok.ToString)
                    node.Token.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type = TokenType.BRACKETCLOSE) Then
                        GoTo Label_02DF
                    End If
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.BRACKETCLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                End If
            End If
            Return
Label_02DF:
            If (Me.scanner.LookAhead(New TokenType() {TokenType.UNARYOPER}).Type = TokenType.UNARYOPER) Then
                tok = Me.scanner.Scan(New TokenType() {TokenType.UNARYOPER})
                n = node.CreateNode(tok, tok.ToString)
                node.Token.UpdateRange(tok)
                node.Nodes.Add(n)
                If (tok.Type <> TokenType.UNARYOPER) Then
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.UNARYOPER.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                    Return
                End If
            End If
            parent.Token.UpdateRange(node.Token)
        End Sub


        ' Fields
        Private scanner As Scanner
        Private tree As ParseTree
    End Class
End Namespace

