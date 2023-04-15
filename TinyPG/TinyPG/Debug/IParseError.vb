Imports System

Namespace TinyPG.Debug
    Public Interface IParseError
        ' Properties
        ReadOnly Property Code As Integer

        ReadOnly Property Column As Integer

        ReadOnly Property Length As Integer

        ReadOnly Property Line As Integer

        ReadOnly Property Message As String

        ReadOnly Property Position As Integer

    End Interface
End Namespace

