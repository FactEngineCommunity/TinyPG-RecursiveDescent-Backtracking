Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace TinyPG.Highlighter
    <Serializable> _
    Public Class ParseTree
        Inherits ParseNode
        ' Methods
        Public Sub New()
            MyBase.New(New Token, "ParseTree")
            MyBase.TokenField.Type = TokenType.Start
            MyBase.TokenField.Text = "Root"
            Me.Errors = New ParseErrors
        End Sub

        Public Shadows Function Eval(ByVal ParamArray paramlist As Object()) As Object
            Return MyBase.Nodes.Item(0).Eval(Me, paramlist)
        End Function

        Private Sub PrintNode(ByVal sb As StringBuilder, ByVal node As ParseNode, ByVal indent As Integer)
            Dim space As String = "".PadLeft(indent, " "c)
            sb.Append(space)
            sb.AppendLine(node.Text)
            Dim n As ParseNode
            For Each n In node.Nodes
                Me.PrintNode(sb, n, (indent + 2))
            Next
        End Sub

        Public Function PrintTree() As String
            Dim sb As New StringBuilder
            Dim indent As Integer = 0
            Me.PrintNode(sb, Me, indent)
            Return sb.ToString
        End Function


        ' Fields
        Public Errors As ParseErrors
        Public Skipped As List(Of Token)
    End Class
End Namespace

