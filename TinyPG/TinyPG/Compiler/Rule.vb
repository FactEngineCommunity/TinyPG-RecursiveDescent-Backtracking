Imports System

Namespace TinyPG.Compiler
    Public Class Rule
        ' Methods
        Public Sub New()
            Me.New(Nothing, RuleType.Choice)
        End Sub

        Public Sub New(ByVal type As RuleType)
            Me.New(Nothing, type)
        End Sub

        Public Sub New(ByVal s As Symbol)
            Me.New(s, If(TypeOf s Is TerminalSymbol, RuleType.Terminal, RuleType.NonTerminal))
        End Sub

        Public Sub New(ByVal s As Symbol, ByVal type As RuleType)
            Me.Type = type
            Me.Symbol = s
            Me.Rules = New Rules
        End Sub

        Friend Function DetermineFirstTerminals(ByVal FirstTerminals As Symbols) As Boolean
            Return Me.DetermineFirstTerminals(FirstTerminals, 0)
        End Function

        Friend Function DetermineFirstTerminals(ByVal FirstTerminals As Symbols, ByVal index As Integer) As Boolean
            Dim containsEmpty As Boolean = False
            Select Case Me.Type
                Case RuleType.Terminal
                    If (Not Me.Symbol Is Nothing) Then
                        If Not FirstTerminals.Exists(Me.Symbol) Then
                            FirstTerminals.Add(Me.Symbol)
                        Else
                            Console.WriteLine("throw new Exception(""Terminal already exists"");")
                        End If
                        Return containsEmpty
                    End If
                    Return True
                Case RuleType.NonTerminal
                    If (Not Me.Symbol Is Nothing) Then
                        Dim nts As NonTerminalSymbol = TryCast(Me.Symbol,NonTerminalSymbol)
                        containsEmpty = nts.DetermineFirstTerminals
                        Dim t As TerminalSymbol
                        For Each t In nts.FirstTerminals
                            If Not FirstTerminals.Exists(t) Then
                                FirstTerminals.Add(t)
                            Else
                                Console.WriteLine("throw new Exception(""Terminal already exists"");")
                            End If
                        Next
                        Return containsEmpty
                    End If
                    Return True
                Case RuleType.Choice
                    Dim r As Rule
                    For Each r In Me.Rules
                        containsEmpty = (containsEmpty Or r.DetermineFirstTerminals(FirstTerminals))
                    Next
                    Return containsEmpty
                Case RuleType.Concat
                    Dim i As Integer
                    For i = index To Me.Rules.Count - 1
                        containsEmpty = Me.Rules.Item(i).DetermineFirstTerminals(FirstTerminals)
                        If Not containsEmpty Then
                            Exit For
                        End If
                    Next i
                    Exit Select
                Case RuleType.Option, RuleType.ZeroOrMore
                    containsEmpty = True
                    Dim r As Rule
                    For Each r In Me.Rules
                        containsEmpty = (containsEmpty Or r.DetermineFirstTerminals(FirstTerminals))
                        If Not containsEmpty Then
                            Return containsEmpty
                        End If
                    Next
                    Return containsEmpty
                Case RuleType.OneOrMore
                    Dim r As Rule
                    For Each r In Me.Rules
                        containsEmpty = r.DetermineFirstTerminals(FirstTerminals)
                        If Not containsEmpty Then
                            Return containsEmpty
                        End If
                    Next
                    Return containsEmpty
                Case Else
                    Throw New NotImplementedException
            End Select
            Dim t1 As TerminalSymbol
            For Each t1 In FirstTerminals
                t1.Rule = Me
            Next
            Return containsEmpty
        End Function

        Public Sub DetermineProductionSymbols(ByVal symbols As Symbols)
            If ((Me.Type = RuleType.Terminal) OrElse (Me.Type = RuleType.NonTerminal)) Then
                symbols.Add(Me.Symbol)
            Else
                Dim rule As Rule
                For Each rule In Me.Rules
                    rule.DetermineProductionSymbols(symbols)
                Next
            End If
        End Sub

        Public Function GetFirstTerminals() As Symbols
            Dim visited As New Symbols
            Dim FirstTerminals As New Symbols
            Me.DetermineFirstTerminals(FirstTerminals)
            Return FirstTerminals
        End Function

        Public Function PrintRule() As String
            Dim r As String = ""
            Select Case Me.Type
                Case RuleType.Terminal, RuleType.NonTerminal
                    If (Not Me.Symbol Is Nothing) Then
                        r = Me.Symbol.Name
                    End If
                    Return r
                Case RuleType.Choice
                    r = (r & "(")
                    Dim rule As Rule
                    For Each rule In Me.Rules
                        If (r.Length > 1) Then
                            r = (r & " | ")
                        End If
                        r = (r & rule.PrintRule)
                    Next
                    r = (r & ")")
                    If (Me.Rules.Count < 1) Then
                        r = (r & " <- WARNING: ChoiceRule contains no subrules")
                    End If
                    Return r
                Case RuleType.Concat
                    Dim rule As Rule
                    For Each rule In Me.Rules
                        r = (r & rule.PrintRule & " ")
                    Next
                    If (Me.Rules.Count < 1) Then
                        r = (r & " <- WARNING: ConcatRule contains no subrules")
                    End If
                    Return r
                Case RuleType.Option
                    If (Me.Rules.Count >= 1) Then
                        r = (r & "(" & Me.Rules.Item(0).PrintRule & ")?")
                    End If
                    If (Me.Rules.Count > 1) Then
                        r = (r & " <- WARNING: OptionRule contains more than 1 subrule")
                    End If
                    If (Me.Rules.Count < 1) Then
                        r = (r & " <- WARNING: OptionRule contains no subrule")
                    End If
                    Return r
                Case RuleType.ZeroOrMore
                    If (Me.Rules.Count >= 1) Then
                        r = (r & "(" & Me.Rules.Item(0).PrintRule & ")*")
                    End If
                    If (Me.Rules.Count > 1) Then
                        r = (r & " <- WARNING: ZeroOrMoreRule contains more than 1 subrule")
                    End If
                    If (Me.Rules.Count < 1) Then
                        r = (r & " <- WARNING: ZeroOrMoreRule contains no subrule")
                    End If
                    Return r
                Case RuleType.OneOrMore
                    If (Me.Rules.Count >= 1) Then
                        r = (r & "(" & Me.Rules.Item(0).PrintRule & ")+")
                    End If
                    If (Me.Rules.Count > 1) Then
                        r = (r & " <- WARNING: OneOrMoreRule contains more than 1 subrule")
                    End If
                    If (Me.Rules.Count < 1) Then
                        r = (r & " <- WARNING: OneOrMoreRule contains no subrule")
                    End If
                    Return r
            End Select
            Return Me.Symbol.Name
        End Function


        ' Fields
        Public Rules As Rules
        Public Symbol As Symbol
        Public Type As RuleType
    End Class
End Namespace

