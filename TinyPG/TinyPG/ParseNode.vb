Imports System
Imports System.Collections.Generic
Imports System.Xml.Serialization

Namespace TinyPG
    <Serializable(), XmlInclude(GetType(ParseTree))> _
    Public Class ParseNode

        Protected nodesField As List(Of ParseNode)
        <XmlIgnore()> _
        Public Parent As ParseNode
        Protected textField As String
        Public Token As Token

        ' Methods
        Protected Sub New(ByVal token As Token, ByVal [text] As String)
            Me.Token = token
            Me.Text = [text]
            Me.nodesField = New List(Of ParseNode)
        End Sub

        Public Overridable Function CreateNode(ByVal token As Token, ByVal [text] As String) As ParseNode
            Return New ParseNode(token, [text]) With { _
                .Parent = Me _
            }
        End Function

        Friend Function Eval(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Select Case Me.Token.Type
                Case TokenType.Start
                    Return Me.EvalStart(tree, paramlist)
                Case TokenType.Directive
                    Return Me.EvalDirective(tree, paramlist)
                Case TokenType.NameValue
                    Return Me.EvalNameValue(tree, paramlist)
                Case TokenType.ExtProduction
                    Return Me.EvalExtProduction(tree, paramlist)
                Case TokenType.Attribute
                    Return Me.EvalAttribute(tree, paramlist)
                Case TokenType.Params
                    Return Me.EvalParams(tree, paramlist)
                Case TokenType.Param
                    Return Me.EvalParam(tree, paramlist)
                Case TokenType.Production
                    Return Me.EvalProduction(tree, paramlist)
                Case TokenType.Rule
                    Return Me.EvalRule(tree, paramlist)
                Case TokenType.Subrule
                    Return Me.EvalSubrule(tree, paramlist)
                Case TokenType.ConcatRule
                    Return Me.EvalConcatRule(tree, paramlist)
                Case TokenType.Symbol
                    Return Me.EvalSymbol(tree, paramlist)
            End Select
            Return Me.Token.Text
        End Function

        Protected Overridable Function EvalAttribute(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalConcatRule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalDirective(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalExtProduction(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalNameValue(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalParam(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalParams(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalProduction(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalRule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalStart(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Return "Could not interpret input; no semantics implemented."
        End Function

        Protected Overridable Function EvalSubrule(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Overridable Function EvalSymbol(ByVal tree As ParseTree, ByVal ParamArray paramlist As Object()) As Object
            Throw New NotImplementedException
        End Function

        Protected Function GetValue(ByVal tree As ParseTree, ByVal type As TokenType, ByRef index As Integer) As Object
            If (index >= 0) Then
                Dim node As ParseNode
                For Each node In Me.Nodes
                    If (node.Token.Type = type) Then
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

        <XmlIgnore()> _
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

