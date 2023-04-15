Imports System
Imports System.Collections.Generic

Namespace TinyPG.Compiler
    Public Class Symbols
        Inherits List(Of Symbol)
        ' Methods
        Public Shadows Function Exists(ByVal symbol As Symbol) As Boolean
            Return MyBase.Exists(Function(item) item.Name = symbol.Name)
        End Function

        Public Shadows Function Find(ByVal Name As String) As Symbol
            Return MyBase.Find(Function(item) ((Not item Is Nothing) AndAlso (item.Name = Name)))
        End Function

    End Class
End Namespace

