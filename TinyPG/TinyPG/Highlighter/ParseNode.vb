Imports System
Imports System.Collections.Generic
Imports System.Xml.Serialization

Namespace TinyPG.Highlighter
    <Serializable, XmlInclude(GetType(ParseTree))> _
    Public Class ParseNode

        ' Fields
        Protected nodesField As List(Of ParseNode)
        <XmlIgnore()> _
        Public ParentField As ParseNode
        Protected textField As String
        Public TokenField As Token

        ' Methods
        Protected Sub New(ByVal token As Token, ByVal [text] As String)
            Me.TokenField = token
            Me.textField = [text]
            Me.nodesField = New List(Of ParseNode)
        End Sub

        Public Overridable Function CreateNode(ByVal token As Token, ByVal [text] As String) As ParseNode
            Return New ParseNode(token, [text]) With { _
                .ParentField = Me _
            }
        End Function

        Friend Function Eval(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Select Case Me.TokenField.Type
                Case TokenType.Start
                    Return Me.EvalStart(tree, paramlist)
                Case TokenType.CommentBlock
                    Return Me.EvalCommentBlock(tree, paramlist)
                Case TokenType.DirectiveBlock
                    Return Me.EvalDirectiveBlock(tree, paramlist)
                Case TokenType.GrammarBlock
                    Return Me.EvalGrammarBlock(tree, paramlist)
                Case TokenType.AttributeBlock
                    Return Me.EvalAttributeBlock(tree, paramlist)
                Case TokenType.CodeBlock
                    Return Me.EvalCodeBlock(tree, paramlist)
            End Select
            Return Me.TokenField.Text
        End Function

        Protected Overridable Function EvalAttributeBlock(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalCodeBlock(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalCommentBlock(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalDirectiveBlock(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalGrammarBlock(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalStart(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Return "Could not interpret input; no semantics implemented."
        End Function

        Protected Function GetValue(ByVal tree As ParseTree, ByVal type As TokenType, ByRef index As Integer) As Object
            If (index >= 0) Then
                Dim node As ParseNode
                For Each node In Me.nodes
                    If (node.TokenField.Type = type) Then
                        index -= 1
                        If (index < 0) Then
                            Return node.Eval(tree, New Object(0 - 1) {})
                        End If
                    End If
                Next
            End If
            Return Nothing
        End Function


        ' Properties
        Public ReadOnly Property Nodes As List(Of ParseNode)
            Get
                Return Me.nodesField
            End Get
        End Property

        <XmlIgnore> _
        Public Property [Text] As String
            Get
                Return Me.textField
            End Get
            Set(ByVal value As String)
                Me.textField = value
            End Set
        End Property
    End Class
End Namespace

