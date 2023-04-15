Imports System

Namespace TinyPG.Debug
    Public Interface IToken
        ' Methods
        Function ToString() As String

        ' Properties
        Property EndPos As Integer

        ReadOnly Property Length As Integer

        Property StartPos As Integer

        Property [Text] As String

    End Interface
End Namespace

