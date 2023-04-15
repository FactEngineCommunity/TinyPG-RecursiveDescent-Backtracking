Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler
Imports Microsoft.VisualBasic

Namespace TinyPG.CodeGenerators.VBNet
    Public Class ParseTreeGenerator
        Implements ICodeGenerator
        ' Methods
        Friend Sub New()
        End Sub

        Private Function FormatCodeBlock(ByVal nts As NonTerminalSymbol) As String
            Dim codeblock As String = nts.CodeBlock
            If (nts Is Nothing) Then
                Return ""
            End If
            Dim var As New Regex("\$(?<var>[a-zA-Z_0-9]+)(\[(?<index>[^]]+)\])?", RegexOptions.Compiled)
            Dim symbols As Symbols = nts.DetermineProductionSymbols
            Dim match As Match = var.Match(codeblock)
            Do While match.Success
                Dim s As Symbol = symbols.Find(match.Groups.Item("var").Value)
                If (s Is Nothing) Then
                    Exit Do
                End If
                Dim indexer As String = "0"
                If (match.Groups.Item("index").Value.Length > 0) Then
                    indexer = match.Groups.Item("index").Value
                End If
                Dim replacement As String = String.Concat(New String() {"Me.GetValue(tree, TokenType.", s.Name, ", ", indexer, ")"})
                codeblock = (codeblock.Substring(0, match.Captures.Item(0).Index) & replacement & codeblock.Substring((match.Captures.Item(0).Index + match.Captures.Item(0).Length)))
                match = var.Match(codeblock)
            Loop
            Return ("            " & codeblock.Replace(ChrW(10), ChrW(13) & ChrW(10) & "        "))
        End Function

        Public Function Generate(ByVal Grammar As Grammar, ByVal Debug As Boolean) As String Implements ICodeGenerator.Generate
            If String.IsNullOrEmpty(Grammar.GetTemplatePath) Then
                Return Nothing
            End If
            Dim parsetree As String = File.ReadAllText((Grammar.GetTemplatePath & Me.FileName))
            Dim evalsymbols As New StringBuilder
            Dim evalmethods As New StringBuilder
            Dim s As Symbol
            For Each s In Grammar.GetNonTerminals
                evalsymbols.AppendLine(("                Case TokenType." & s.Name))
                evalsymbols.AppendLine(("                    Value = Eval" & s.Name & "(tree, paramlist)"))
                evalsymbols.AppendLine("                    Exit Select")
                evalmethods.AppendLine(("        Protected Overridable Function Eval" & s.Name & "(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object"))
                If (Not s.CodeBlock Is Nothing) Then
                    evalmethods.AppendLine(Me.FormatCodeBlock(TryCast(s, NonTerminalSymbol)))
                ElseIf (s.Name = "Start") Then
                    evalmethods.AppendLine("            Return ""Could not interpret input; no semantics implemented.""")
                Else
                    evalmethods.AppendLine("            Throw New NotImplementedException()")
                End If
                evalmethods.AppendLine("        End Function" & ChrW(13) & ChrW(10))
            Next
            If Debug Then
                parsetree = parsetree.Replace("<%Imports%>", "Imports TinyPG.Debug").Replace("<%Namespace%>", "TinyPG.Debug").Replace("<%IParseTree%>", ChrW(13) & ChrW(10) & "        Implements IParseTree").Replace("<%IParseNode%>", ChrW(13) & ChrW(10) & "        Implements IParseNode" & ChrW(13) & ChrW(10)).Replace("<%ParseError%>", ChrW(13) & ChrW(10) & "        Implements IParseError" & ChrW(13) & ChrW(10)).Replace("<%ParseErrors%>", "List(Of IParseError)")
                Dim itoken As String = "        Public ReadOnly Property IToken() As IToken Implements IParseNode.IToken" & ChrW(13) & ChrW(10) & "            Get" & ChrW(13) & ChrW(10) & "                Return DirectCast(Token, IToken)" & ChrW(13) & ChrW(10) & "            End Get" & ChrW(13) & ChrW(10) & "        End Property" & ChrW(13) & ChrW(10)
                parsetree = parsetree.Replace("<%ITokenGet%>", itoken).Replace("<%ImplementsIParseTreePrintTree%>", " Implements IParseTree.PrintTree").Replace("<%ImplementsIParseTreeEval%>", " Implements IParseTree.Eval").Replace("<%ImplementsIParseErrorCode%>", " Implements IParseError.Code").Replace("<%ImplementsIParseErrorLine%>", " Implements IParseError.Line").Replace("<%ImplementsIParseErrorColumn%>", " Implements IParseError.Column").Replace("<%ImplementsIParseErrorPosition%>", " Implements IParseError.Position").Replace("<%ImplementsIParseErrorLength%>", " Implements IParseError.Length").Replace("<%ImplementsIParseErrorMessage%>", " Implements IParseError.Message")
                Dim inodes As String = "        Public Shared Function Node2INode(ByVal node As ParseNode) As IParseNode" & ChrW(13) & ChrW(10) & "            Return DirectCast(node, IParseNode)" & ChrW(13) & ChrW(10) & "        End Function" & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10) & "        Public ReadOnly Property INodes() As List(Of IParseNode) Implements IParseNode.INodes" & ChrW(13) & ChrW(10) & "            Get" & ChrW(13) & ChrW(10) & "                Return Nodes.ConvertAll(Of IParseNode)(New Converter(Of ParseNode, IParseNode)(AddressOf Node2INode))" & ChrW(13) & ChrW(10) & "            End Get" & ChrW(13) & ChrW(10) & "        End Property" & ChrW(13) & ChrW(10)
                parsetree = parsetree.Replace("<%INodesGet%>", inodes).Replace("<%ImplementsIParseNodeText%>", " Implements IParseNode.Text")
            Else
                parsetree = parsetree.Replace("<%Imports%>", "").Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace")).Replace("<%ParseError%>", "").Replace("<%ParseErrors%>", "List(Of ParseError)").Replace("<%IParseTree%>", "").Replace("<%IParseNode%>", "").Replace("<%ITokenGet%>", "").Replace("<%INodesGet%>", "").Replace("<%ImplementsIParseTreePrintTree%>", "").Replace("<%ImplementsIParseTreeEval%>", "").Replace("<%ImplementsIParseErrorCode%>", "").Replace("<%ImplementsIParseErrorLine%>", "").Replace("<%ImplementsIParseErrorColumn%>", "").Replace("<%ImplementsIParseErrorPosition%>", "").Replace("<%ImplementsIParseErrorLength%>", "").Replace("<%ImplementsIParseErrorMessage%>", "").Replace("<%ImplementsIParseNodeText%>", "")
            End If
            Return parsetree.Replace("<%EvalSymbols%>", evalsymbols.ToString).Replace("<%VirtualEvalMethods%>", evalmethods.ToString)
        End Function


        ' Properties
        Public ReadOnly Property FileName As String Implements ICodeGenerator.FileName
            Get
                Return "ParseTree.vb"
            End Get
        End Property

    End Class
End Namespace

