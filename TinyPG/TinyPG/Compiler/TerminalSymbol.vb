Imports System
Imports System.Text
Imports System.Text.RegularExpressions

Namespace TinyPG.Compiler
    Public Class TerminalSymbol
        Inherits Symbol
        ' Methods
        Public Sub New()
            Me.New(("Terminal_" & ++Symbol.counter), "")
        End Sub

        Public Sub New(ByVal name As String)
            Me.New(name, "")
        End Sub

        Public Sub New(ByVal name As String, ByVal pattern As String)
            MyBase.Name = name
            Me.Expression = New Regex(pattern, RegexOptions.Compiled)
        End Sub

        Public Sub New(ByVal name As String, ByVal expression As Regex)
            MyBase.Name = name
            Me.Expression = expression
        End Sub

        Public Overrides Function PrintProduction() As String
            Return Helper.Outline(MyBase.Name, 0, (" -> " & Me.Expression.ToString & ";"), 4)
        End Function


        ' Fields
        Public Expression As Regex
    End Class
End Namespace

