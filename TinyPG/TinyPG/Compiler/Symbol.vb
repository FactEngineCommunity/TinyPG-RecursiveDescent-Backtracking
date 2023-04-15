Imports System

Namespace TinyPG.Compiler
    Public MustInherit Class Symbol
        ' Methods
        Protected Sub New()
        End Sub

        Public MustOverride Function PrintProduction() As String


        ' Fields
        Public Attributes As SymbolAttributes = New SymbolAttributes
        Public CodeBlock As String
        Protected Shared counter As Integer = 0
        Public Name As String
        Public Rule As Rule
    End Class
End Namespace

