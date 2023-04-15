Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler
Imports Microsoft.VisualBasic

Namespace TinyPG.CodeGenerators.CSharp
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
                Dim replacement As String = String.Concat(New String() { "this.GetValue(tree, TokenType.", s.Name, ", ", indexer, ")" })
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
                evalsymbols.AppendLine(("                case TokenType." & s.Name & ":"))
                evalsymbols.AppendLine(("                    Value = Eval" & s.Name & "(tree, paramlist);"))
                evalsymbols.AppendLine("                    break;")
                evalmethods.AppendLine(("        protected virtual object Eval" & s.Name & "(ParseTree tree, params object[] paramlist)"))
                evalmethods.AppendLine("        {")
                If (Not s.CodeBlock Is Nothing) Then
                    evalmethods.AppendLine(Me.FormatCodeBlock(TryCast(s, NonTerminalSymbol)))
                ElseIf (s.Name = "Start") Then
                    evalmethods.AppendLine("            return ""Could not interpret input; no semantics implemented."";")
                Else
                    evalmethods.AppendLine("            throw new NotImplementedException();")
                End If
                evalmethods.AppendLine("        }" & ChrW(13) & ChrW(10))
            Next
            If Debug Then
                parsetree = parsetree.Replace("<%Namespace%>", "TinyPG.Debug").Replace("<%ParseError%>", " : TinyPG.Debug.IParseError").Replace("<%ParseErrors%>", "List<TinyPG.Debug.IParseError>").Replace("<%IParseTree%>", ", TinyPG.Debug.IParseTree").Replace("<%IParseNode%>", " : TinyPG.Debug.IParseNode").Replace("<%ITokenGet%>", "public IToken IToken { get {return (IToken)Token;} }")
                Dim inodes As String = "public List<IParseNode> INodes {get { return nodes.ConvertAll<IParseNode>( new Converter<ParseNode, IParseNode>( delegate(ParseNode n) { return (IParseNode)n; })); }}" & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10)
                parsetree = parsetree.Replace("<%INodesGet%>", inodes)
            Else
                parsetree = parsetree.Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace")).Replace("<%ParseError%>", "").Replace("<%ParseErrors%>", "List<ParseError>").Replace("<%IParseTree%>", "").Replace("<%IParseNode%>", "").Replace("<%ITokenGet%>", "").Replace("<%INodesGet%>", "")
            End If
            Return parsetree.Replace("<%EvalSymbols%>", evalsymbols.ToString).Replace("<%VirtualEvalMethods%>", evalmethods.ToString)
        End Function


        ' Properties
        Public ReadOnly Property FileName As String Implements ICodeGenerator.FileName
            Get
                Return "ParseTree.cs"
            End Get
        End Property

    End Class
End Namespace

