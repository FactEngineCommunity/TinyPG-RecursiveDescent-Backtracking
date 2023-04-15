Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Namespace TinyPG
    Public Class Scanner
        ' Methods
        Public Sub New()
            Me.SkipList.Add(TokenType.WHITESPACE)
            Me.SkipList.Add(TokenType.COMMENTLINE)
            Me.SkipList.Add(TokenType.COMMENTBLOCK)
            Dim regex As New Regex("\(", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.BRACKETOPEN, regex)
            Me.Tokens.Add(TokenType.BRACKETOPEN)
            regex = New Regex("\)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.BRACKETCLOSE, regex)
            Me.Tokens.Add(TokenType.BRACKETCLOSE)
            regex = New Regex("\{[^\}]*\}([^};][^}]*\}+)*;", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.CODEBLOCK, regex)
            Me.Tokens.Add(TokenType.CODEBLOCK)
            regex = New Regex(",", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.COMMA, regex)
            Me.Tokens.Add(TokenType.COMMA)
            regex = New Regex("\[", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.SQUAREOPEN, regex)
            Me.Tokens.Add(TokenType.SQUAREOPEN)
            regex = New Regex("\]", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.SQUARECLOSE, regex)
            Me.Tokens.Add(TokenType.SQUARECLOSE)
            regex = New Regex("=", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ASSIGN, regex)
            Me.Tokens.Add(TokenType.ASSIGN)
            regex = New Regex("\|", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.PIPE, regex)
            Me.Tokens.Add(TokenType.PIPE)
            regex = New Regex(";", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.SEMICOLON, regex)
            Me.Tokens.Add(TokenType.SEMICOLON)
            regex = New Regex("(\*|\+|\?)", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.UNARYOPER, regex)
            Me.Tokens.Add(TokenType.UNARYOPER)
            regex = New Regex("[a-zA-Z_][a-zA-Z0-9_]*", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.IDENTIFIER, regex)
            Me.Tokens.Add(TokenType.IDENTIFIER)
            regex = New Regex("[0-9]+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.INTEGER, regex)
            Me.Tokens.Add(TokenType.INTEGER)
            regex = New Regex("[0-9]*\.[0-9]+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DOUBLE, regex)
            Me.Tokens.Add(TokenType.DOUBLE)
            regex = New Regex("(0x[0-9a-fA-F]{6})", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.HEX, regex)
            Me.Tokens.Add(TokenType.HEX)
            regex = New Regex("->", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.ARROW, regex)
            Me.Tokens.Add(TokenType.ARROW)
            regex = New Regex("<%\s*@", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVEOPEN, regex)
            Me.Tokens.Add(TokenType.DIRECTIVEOPEN)
            regex = New Regex("%>", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.DIRECTIVECLOSE, regex)
            Me.Tokens.Add(TokenType.DIRECTIVECLOSE)
            regex = New Regex("^$", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.EOF, regex)
            Me.Tokens.Add(TokenType.EOF)
            regex = New Regex("@?\""(\""\""|[^\""])*\""", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.STRING, regex)
            Me.Tokens.Add(TokenType.STRING)
            regex = New Regex("\s+", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.WHITESPACE, regex)
            Me.Tokens.Add(TokenType.WHITESPACE)
            regex = New Regex("//[^\n]*\n?", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.COMMENTLINE, regex)
            Me.Tokens.Add(TokenType.COMMENTLINE)
            regex = New Regex("/\*[^*]*\*+(?:[^/*][^*]*\*+)*/", RegexOptions.Compiled)
            Me.Patterns.Add(TokenType.COMMENTBLOCK, regex)
            Me.Tokens.Add(TokenType.COMMENTBLOCK)
        End Sub

        Public Function GetToken(ByVal type As TokenType) As Token
            Return New Token(Me.StartPos, Me.EndPos) With { _
                .Type = type _
            }
        End Function

        Public Sub Init(ByVal input As String)
            Me.Input = input
            Me.StartPos = 0
            Me.EndPos = 0
            Me.CurrentLine = 0
            Me.CurrentColumn = 0
            Me.CurrentPosition = 0
            Me.LookAheadToken = Nothing
        End Sub

        Public Function LookAhead(ByVal ParamArray expectedtokens As TokenType()) As Token
            Dim scantokens As List(Of TokenType)
            Dim startpos As Integer = Me.StartPos
            Dim tok As Token = Nothing
            If (((Not Me.LookAheadToken Is Nothing) AndAlso (Me.LookAheadToken.Type <> TokenType._UNDETERMINED_)) AndAlso (Me.LookAheadToken.Type <> TokenType._NONE_)) Then
                Return Me.LookAheadToken
            End If
            If (expectedtokens.Length = 0) Then
                scantokens = Me.Tokens
            Else
                scantokens = New List(Of TokenType)(expectedtokens)
                scantokens.AddRange(Me.SkipList)
            End If
            Do
                Dim len As Integer = -1
                Dim index As TokenType = DirectCast(&H7FFFFFFF, TokenType)
                Dim input As String = Me.Input.Substring(startpos)
                tok = New Token(startpos, Me.EndPos)
                Dim i As Integer
                For i = 0 To scantokens.Count - 1
                    Dim m As Match = Me.Patterns.Item(scantokens.Item(i)).Match(input)
                    If ((m.Success AndAlso (m.Index = 0)) AndAlso ((m.Length > len) OrElse ((DirectCast(scantokens.Item(i), TokenType) < index) AndAlso (m.Length = len)))) Then
                        len = m.Length
                        index = scantokens.Item(i)
                    End If
                Next i
                If ((index >= TokenType._NONE_) AndAlso (len >= 0)) Then
                    tok.EndPos = (startpos + len)
                    tok.Text = Me.Input.Substring(tok.StartPos, len)
                    tok.Type = index
                ElseIf (tok.StartPos < (tok.EndPos - 1)) Then
                    tok.Text = Me.Input.Substring(tok.StartPos, 1)
                End If
                If Me.SkipList.Contains(tok.Type) Then
                    startpos = tok.EndPos
                    Me.Skipped.Add(tok)
                Else
                    tok.Skipped = Me.Skipped
                    Me.Skipped = New List(Of Token)
                End If
            Loop While Me.SkipList.Contains(tok.Type)
            Me.LookAheadToken = tok
            Return tok
        End Function

        Public Function Scan(ByVal ParamArray expectedtokens As TokenType()) As Token
            Dim tok As Token = Me.LookAhead(expectedtokens)
            Me.LookAheadToken = Nothing
            Me.StartPos = tok.EndPos
            Me.EndPos = tok.EndPos
            Return tok
        End Function


        ' Fields
        Public CurrentColumn As Integer
        Public CurrentLine As Integer
        Public CurrentPosition As Integer
        Public EndPos As Integer = 0
        Public Input As String
        Private LookAheadToken As Token = Nothing
        Public Patterns As Dictionary(Of TokenType, Regex) = New Dictionary(Of TokenType, Regex)
        Private SkipList As List(Of TokenType) = New List(Of TokenType)
        Public Skipped As List(Of Token) = New List(Of Token)
        Public StartPos As Integer = 0
        Private Tokens As List(Of TokenType) = New List(Of TokenType)
    End Class
End Namespace

