Imports System
Imports System.Collections.Generic

Namespace TinyPG.Debug
    Public Interface IParseNode
        ' Properties
        ReadOnly Property INodes As List(Of IParseNode)

        ReadOnly Property IToken As IToken

        Property [Text] As String

    End Interface
End Namespace

