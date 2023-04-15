Imports System
Imports Microsoft.VisualBasic

Namespace TinyPG.Highlighter
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

        Private Sub ParseAttributeBlock(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.AttributeBlock), "AttributeBlock")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() { TokenType.ATTRIBUTEOPEN })
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.TokenField.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.ATTRIBUTEOPEN) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ATTRIBUTEOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.ATTRIBUTEKEYWORD, TokenType.ATTRIBUTENONKEYWORD, TokenType.ATTRIBUTESYMBOL })
                Do While (((tok.Type = TokenType.ATTRIBUTEKEYWORD) OrElse (tok.Type = TokenType.ATTRIBUTENONKEYWORD)) OrElse (tok.Type = TokenType.ATTRIBUTESYMBOL))
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.ATTRIBUTEKEYWORD, TokenType.ATTRIBUTENONKEYWORD, TokenType.ATTRIBUTESYMBOL })
                    Select Case tok.Type
                        Case TokenType.ATTRIBUTESYMBOL
                            tok = Me.scanner.Scan(New TokenType() { TokenType.ATTRIBUTESYMBOL })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.ATTRIBUTESYMBOL) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ATTRIBUTESYMBOL.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.ATTRIBUTEKEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.ATTRIBUTEKEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.ATTRIBUTEKEYWORD) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ATTRIBUTEKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.ATTRIBUTENONKEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.ATTRIBUTENONKEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.ATTRIBUTENONKEYWORD) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ATTRIBUTENONKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case Else
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Exit Select
                    End Select
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.ATTRIBUTEKEYWORD, TokenType.ATTRIBUTENONKEYWORD, TokenType.ATTRIBUTESYMBOL })
                Loop
                If (Me.scanner.LookAhead(New TokenType() { TokenType.ATTRIBUTECLOSE }).Type = TokenType.ATTRIBUTECLOSE) Then
                    tok = Me.scanner.Scan(New TokenType() { TokenType.ATTRIBUTECLOSE })
                    n = node.CreateNode(tok, tok.ToString)
                    node.TokenField.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.ATTRIBUTECLOSE) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.ATTRIBUTECLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    End If
                End If
                parent.TokenField.UpdateRange(node.TokenField)
            End If
        End Sub

        Private Sub ParseCodeBlock(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.CodeBlock), "CodeBlock")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() { TokenType.CODEBLOCKOPEN })
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.TokenField.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.CODEBLOCKOPEN) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.CODEBLOCKOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.DOTNET_COMMENTLINE, TokenType.DOTNET_COMMENTBLOCK, TokenType.DOTNET_TYPES, TokenType.DOTNET_KEYWORD, TokenType.DOTNET_SYMBOL, TokenType.DOTNET_STRING, TokenType.DOTNET_NONKEYWORD })
                Do While (((((tok.Type = TokenType.DOTNET_COMMENTLINE) OrElse (tok.Type = TokenType.DOTNET_COMMENTBLOCK)) OrElse ((tok.Type = TokenType.DOTNET_TYPES) OrElse (tok.Type = TokenType.DOTNET_KEYWORD))) OrElse ((tok.Type = TokenType.DOTNET_SYMBOL) OrElse (tok.Type = TokenType.DOTNET_STRING))) OrElse (tok.Type = TokenType.DOTNET_NONKEYWORD))
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.DOTNET_COMMENTLINE, TokenType.DOTNET_COMMENTBLOCK, TokenType.DOTNET_TYPES, TokenType.DOTNET_KEYWORD, TokenType.DOTNET_SYMBOL, TokenType.DOTNET_STRING, TokenType.DOTNET_NONKEYWORD })
                    Select Case tok.Type
                        Case TokenType.DOTNET_KEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_KEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_KEYWORD) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_KEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_TYPES
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_TYPES })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_TYPES) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_TYPES.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_COMMENTLINE
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_COMMENTLINE })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_COMMENTLINE) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_COMMENTLINE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_COMMENTBLOCK
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_COMMENTBLOCK })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_COMMENTBLOCK) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_COMMENTBLOCK.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_SYMBOL
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_SYMBOL })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_SYMBOL) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_SYMBOL.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_NONKEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_NONKEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_NONKEYWORD) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_NONKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DOTNET_STRING
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DOTNET_STRING })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DOTNET_STRING) Then
                                Exit Select
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DOTNET_STRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case Else
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Exit Select
                    End Select
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.DOTNET_COMMENTLINE, TokenType.DOTNET_COMMENTBLOCK, TokenType.DOTNET_TYPES, TokenType.DOTNET_KEYWORD, TokenType.DOTNET_SYMBOL, TokenType.DOTNET_STRING, TokenType.DOTNET_NONKEYWORD })
                Loop
                If (Me.scanner.LookAhead(New TokenType() { TokenType.CODEBLOCKCLOSE }).Type = TokenType.CODEBLOCKCLOSE) Then
                    tok = Me.scanner.Scan(New TokenType() { TokenType.CODEBLOCKCLOSE })
                    n = node.CreateNode(tok, tok.ToString)
                    node.TokenField.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.CODEBLOCKCLOSE) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.CODEBLOCKCLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    End If
                End If
                parent.TokenField.UpdateRange(node.TokenField)
            End If
        End Sub

        Private Sub ParseCommentBlock(ByVal parent As ParseNode)
            Dim tok As Token
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.CommentBlock), "CommentBlock")
            parent.Nodes.Add(node)
            Do
                Dim n As ParseNode
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK })
                Select Case tok.Type
                    Case TokenType.GRAMMARCOMMENTLINE
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARCOMMENTLINE })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARCOMMENTLINE) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARCOMMENTLINE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case TokenType.GRAMMARCOMMENTBLOCK
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARCOMMENTBLOCK })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARCOMMENTBLOCK) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARCOMMENTBLOCK.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case Else
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Exit Select
                End Select
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK })
            Loop While ((tok.Type = TokenType.GRAMMARCOMMENTLINE) OrElse (tok.Type = TokenType.GRAMMARCOMMENTBLOCK))
            parent.TokenField.UpdateRange(node.TokenField)
        End Sub

        Private Sub ParseDirectiveBlock(ByVal parent As ParseNode)
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.DirectiveBlock), "DirectiveBlock")
            parent.Nodes.Add(node)
            Dim tok As Token = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVEOPEN })
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.TokenField.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.DIRECTIVEOPEN) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVEOPEN.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.WHITESPACE, TokenType.DIRECTIVEKEYWORD, TokenType.DIRECTIVESYMBOL, TokenType.DIRECTIVENONKEYWORD, TokenType.DIRECTIVESTRING })
                Do While ((((tok.Type = TokenType.WHITESPACE) OrElse (tok.Type = TokenType.DIRECTIVEKEYWORD)) OrElse ((tok.Type = TokenType.DIRECTIVESYMBOL) OrElse (tok.Type = TokenType.DIRECTIVENONKEYWORD))) OrElse (tok.Type = TokenType.DIRECTIVESTRING))
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.WHITESPACE, TokenType.DIRECTIVEKEYWORD, TokenType.DIRECTIVESYMBOL, TokenType.DIRECTIVENONKEYWORD, TokenType.DIRECTIVESTRING })
                    Select Case tok.Type
                        Case TokenType.WHITESPACE
                            tok = Me.scanner.Scan(New TokenType() { TokenType.WHITESPACE })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.WHITESPACE) Then
                                Continue Do
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.WHITESPACE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DIRECTIVESTRING
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVESTRING })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DIRECTIVESTRING) Then
                                Continue Do
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVESTRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DIRECTIVEKEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVEKEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DIRECTIVEKEYWORD) Then
                                Continue Do
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVEKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DIRECTIVESYMBOL
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVESYMBOL })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DIRECTIVESYMBOL) Then
                                Continue Do
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVESYMBOL.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                        Case TokenType.DIRECTIVENONKEYWORD
                            tok = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVENONKEYWORD })
                            n = node.CreateNode(tok, tok.ToString)
                            node.TokenField.UpdateRange(tok)
                            node.Nodes.Add(n)
                            If (tok.Type = TokenType.DIRECTIVENONKEYWORD) Then
                                Continue Do
                            End If
                            Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVENONKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                            Return
                    End Select
                    Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                    tok = Me.scanner.LookAhead(New TokenType() { TokenType.WHITESPACE, TokenType.DIRECTIVEKEYWORD, TokenType.DIRECTIVESYMBOL, TokenType.DIRECTIVENONKEYWORD, TokenType.DIRECTIVESTRING })
                Loop
                If (Me.scanner.LookAhead(New TokenType() { TokenType.DIRECTIVECLOSE }).Type = TokenType.DIRECTIVECLOSE) Then
                    tok = Me.scanner.Scan(New TokenType() { TokenType.DIRECTIVECLOSE })
                    n = node.CreateNode(tok, tok.ToString)
                    node.TokenField.UpdateRange(tok)
                    node.Nodes.Add(n)
                    If (tok.Type <> TokenType.DIRECTIVECLOSE) Then
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.DIRECTIVECLOSE.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    End If
                End If
                parent.TokenField.UpdateRange(node.TokenField)
            End If
        End Sub

        Private Sub ParseGrammarBlock(ByVal parent As ParseNode)
            Dim tok As Token
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.GrammarBlock), "GrammarBlock")
            parent.Nodes.Add(node)
            Do
                Dim n As ParseNode
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARSTRING, TokenType.GRAMMARARROW, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARKEYWORD, TokenType.GRAMMARSYMBOL })
                Select Case tok.Type
                    Case TokenType.GRAMMARKEYWORD
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARKEYWORD })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARKEYWORD) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case TokenType.GRAMMARARROW
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARARROW })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARARROW) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARARROW.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case TokenType.GRAMMARSYMBOL
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARSYMBOL })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARSYMBOL) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARSYMBOL.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case TokenType.GRAMMARNONKEYWORD
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARNONKEYWORD })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARNONKEYWORD) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARNONKEYWORD.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case TokenType.GRAMMARSTRING
                        tok = Me.scanner.Scan(New TokenType() { TokenType.GRAMMARSTRING })
                        n = node.CreateNode(tok, tok.ToString)
                        node.TokenField.UpdateRange(tok)
                        node.Nodes.Add(n)
                        If (tok.Type = TokenType.GRAMMARSTRING) Then
                            Exit Select
                        End If
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.GRAMMARSTRING.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Return
                    Case Else
                        Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                        Exit Select
                End Select
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARSTRING, TokenType.GRAMMARARROW, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARKEYWORD, TokenType.GRAMMARSYMBOL })
            Loop While ((((tok.Type = TokenType.GRAMMARSTRING) OrElse (tok.Type = TokenType.GRAMMARARROW)) OrElse ((tok.Type = TokenType.GRAMMARNONKEYWORD) OrElse (tok.Type = TokenType.GRAMMARKEYWORD))) OrElse (tok.Type = TokenType.GRAMMARSYMBOL))
            parent.TokenField.UpdateRange(node.TokenField)
        End Sub

        Private Sub ParseStart(ByVal parent As ParseNode)
            Dim tok As Token
            Dim node As ParseNode = parent.CreateNode(Me.scanner.GetToken(TokenType.Start), "Start")
            parent.Nodes.Add(node)
            tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK, TokenType.ATTRIBUTEOPEN, TokenType.GRAMMARSTRING, TokenType.GRAMMARARROW, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARKEYWORD, TokenType.GRAMMARSYMBOL, TokenType.CODEBLOCKOPEN, TokenType.DIRECTIVEOPEN })
            Do While ((((((tok.Type = TokenType.GRAMMARCOMMENTLINE) OrElse (tok.Type = TokenType.GRAMMARCOMMENTBLOCK)) OrElse ((tok.Type = TokenType.ATTRIBUTEOPEN) OrElse (tok.Type = TokenType.GRAMMARSTRING))) OrElse (((tok.Type = TokenType.GRAMMARARROW) OrElse (tok.Type = TokenType.GRAMMARNONKEYWORD)) OrElse ((tok.Type = TokenType.GRAMMARKEYWORD) OrElse (tok.Type = TokenType.GRAMMARSYMBOL)))) OrElse (tok.Type = TokenType.CODEBLOCKOPEN)) OrElse (tok.Type = TokenType.DIRECTIVEOPEN))
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK, TokenType.ATTRIBUTEOPEN, TokenType.GRAMMARSTRING, TokenType.GRAMMARARROW, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARKEYWORD, TokenType.GRAMMARSYMBOL, TokenType.CODEBLOCKOPEN, TokenType.DIRECTIVEOPEN })
                Select Case tok.Type
                    Case TokenType.CODEBLOCKOPEN
                        Me.ParseCodeBlock(node)
                        Continue Do
                    Case TokenType.GRAMMARKEYWORD, TokenType.GRAMMARARROW, TokenType.GRAMMARSYMBOL, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARSTRING
                        Me.ParseGrammarBlock(node)
                        Continue Do
                    Case TokenType.ATTRIBUTEOPEN
                        Me.ParseAttributeBlock(node)
                        Continue Do
                    Case TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK
                        Me.ParseCommentBlock(node)
                        Continue Do
                    Case TokenType.DIRECTIVEOPEN
                        Me.ParseDirectiveBlock(node)
                        Continue Do
                End Select
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found."), 2, 0, tok.StartPos, tok.StartPos, tok.Length))
                tok = Me.scanner.LookAhead(New TokenType() { TokenType.GRAMMARCOMMENTLINE, TokenType.GRAMMARCOMMENTBLOCK, TokenType.ATTRIBUTEOPEN, TokenType.GRAMMARSTRING, TokenType.GRAMMARARROW, TokenType.GRAMMARNONKEYWORD, TokenType.GRAMMARKEYWORD, TokenType.GRAMMARSYMBOL, TokenType.CODEBLOCKOPEN, TokenType.DIRECTIVEOPEN })
            Loop
            tok = Me.scanner.Scan(New TokenType() { TokenType.EOF })
            Dim n As ParseNode = node.CreateNode(tok, tok.ToString)
            node.TokenField.UpdateRange(tok)
            node.Nodes.Add(n)
            If (tok.Type <> TokenType.EOF) Then
                Me.tree.Errors.Add(New ParseError(("Unexpected token '" & tok.Text.Replace(ChrW(10), "") & "' found. Expected " & TokenType.EOF.ToString), &H1001, 0, tok.StartPos, tok.StartPos, tok.Length))
            Else
                parent.TokenField.UpdateRange(node.TokenField)
            End If
        End Sub


        ' Fields
        Private scanner As Scanner
        Private tree As ParseTree
    End Class
End Namespace

