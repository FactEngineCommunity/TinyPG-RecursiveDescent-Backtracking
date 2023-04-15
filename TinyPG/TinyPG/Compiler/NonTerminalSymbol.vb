Imports System
Imports System.Text

Namespace TinyPG.Compiler
    Public Class NonTerminalSymbol
        Inherits Symbol
        ' Methods
        Public Sub New()
            Me.New(("NTS_" & ++Symbol.counter))
        End Sub

        Public Sub New(ByVal name As String)
            Me.FirstTerminals = New Symbols
            Me.Rules = New Rules
            MyBase.Name = name
            Me.containsEmpty = False
            Me.visitCount = 0
        End Sub

        Friend Function DetermineFirstTerminals() As Boolean
            If (Me.visitCount <= 10) Then
                Me.visitCount += 1
                Me.FirstTerminals = New Symbols
                Dim rule As Rule
                For Each rule In Me.Rules
                    Me.containsEmpty = (Me.containsEmpty Or rule.DetermineFirstTerminals(Me.FirstTerminals))
                Next
            End If
            Return Me.containsEmpty
        End Function

        Public Function DetermineProductionSymbols() As Symbols
            Dim symbols As New Symbols
            Dim rule As Rule
            For Each rule In Me.Rules
                rule.DetermineProductionSymbols(symbols)
            Next
            Return symbols
        End Function

        Public Overrides Function PrintProduction() As String
            Dim p As String = ""
            Dim r As Rule
            For Each r In Me.Rules
                p = (p & r.PrintRule & ";")
            Next
            Return Helper.Outline(MyBase.Name, 0, (" -> " & p), 4)
        End Function


        ' Fields
        Private containsEmpty As Boolean
        Public FirstTerminals As Symbols
        Public Rules As Rules
        Private visitCount As Integer
    End Class
End Namespace

