Imports System
Imports System.Globalization
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports Microsoft.VisualBasic

Namespace TinyPG.Compiler
    Public Class Grammar

        Private directivesField As Directives
        Private skipSymbolsField As Symbols
        Private symbolsField As Symbols

        Public Property Directives As Directives
            Get
                Return Me.directivesField
            End Get
            Set(ByVal value As Directives)
                Me.directivesField = value
            End Set
        End Property

        Public Property SkipSymbols As Symbols
            Get
                Return Me.skipSymbolsField
            End Get
            Set(ByVal value As Symbols)
                Me.skipSymbolsField = value
            End Set
        End Property

        Public Property Symbols As Symbols
            Get
                Return Me.symbolsField
            End Get
            Set(ByVal value As Symbols)
                Me.symbolsField = value
            End Set
        End Property

                ' Methods
        Public Sub New()
            Me.Symbols = New Symbols
            Me.SkipSymbols = New Symbols
            Me.Directives = New Directives
        End Sub

        Private Sub DetermineFirsts()
            Dim nts As NonTerminalSymbol
            For Each nts In Me.GetNonTerminals
                nts.DetermineFirstTerminals()
            Next
        End Sub

        Public Function GetNonTerminals() As Symbols
            Dim symbols As New Symbols
            Dim s As Symbol
            For Each s In Me.Symbols
                If TypeOf s Is NonTerminalSymbol Then
                    symbols.Add(s)
                End If
            Next
            Return symbols
        End Function

        Public Function GetOutputPath() As String
            Dim folder As String = (Directory.GetCurrentDirectory & "\")
            Dim pathout As String = Me.Directives.Item("TinyPG").Item("OutputPath")
            If Path.IsPathRooted(pathout) Then
                folder = Path.GetFullPath(pathout)
            Else
                folder = Path.GetFullPath((folder & pathout))
            End If
            Dim dir As New DirectoryInfo((folder & "\"))
            If Not dir.Exists Then
                dir.Create()
            End If
            Return folder
        End Function

        Public Function GetTemplatePath() As String
            Dim folder As String = AppDomain.CurrentDomain.BaseDirectory
            Dim pathout As String = Me.Directives.Item("TinyPG").Item("TemplatePath")
            If Path.IsPathRooted(pathout) Then
                folder = Path.GetFullPath(pathout)
            Else
                folder = Path.GetFullPath((folder & pathout))
            End If
            Dim dir As New DirectoryInfo((folder & "\"))
            If Not dir.Exists Then
                dir.Create()
            End If
            Return folder
        End Function

        Public Function GetTerminals() As Symbols
            Dim symbols As New Symbols
            Dim s As Symbol
            For Each s In Me.Symbols
                If TypeOf s Is TerminalSymbol Then
                    symbols.Add(s)
                End If
            Next
            Return symbols
        End Function

        Public Sub Preprocess()
            Me.SetupDirectives()
            Me.DetermineFirsts()
        End Sub

        Public Function PrintFirsts() As String
            Dim sb As New StringBuilder
            sb.AppendLine(ChrW(13) & ChrW(10) & "/*" & ChrW(13) & ChrW(10) & "First symbols:")
            Dim s As NonTerminalSymbol
            For Each s In Me.GetNonTerminals
                Dim firsts As String = (s.Name & ": ")
                Dim t As TerminalSymbol
                For Each t In s.FirstTerminals
                    firsts = (firsts & t.Name & " "c)
                Next
                sb.AppendLine(firsts)
            Next
            sb.AppendLine(ChrW(13) & ChrW(10) & "Skip symbols: ")
            Dim skips As String = ""
            Dim nts As TerminalSymbol
            For Each nts In Me.SkipSymbols
                skips = (skips & nts.Name & " ")
            Next
            sb.AppendLine(skips)
            sb.AppendLine("*/")
            Return sb.ToString
        End Function

        Public Function PrintGrammar() As String
            Dim sb As New StringBuilder
            sb.AppendLine("//Terminals:")
            Dim s As Symbol
            For Each s In Me.GetTerminals
                If (Not Me.SkipSymbols.Find(s.Name) Is Nothing) Then
                    sb.Append("[Skip] ")
                End If
                sb.AppendLine(s.PrintProduction)
            Next
            sb.AppendLine(ChrW(13) & ChrW(10) & "//Production lines:")
            'Dim s As Symbol
            For Each s In Me.GetNonTerminals
                sb.AppendLine(s.PrintProduction)
            Next
            Return sb.ToString
        End Function

        Private Sub SetupDirectives()
            Dim d As Directive = Me.Directives.Find("TinyPG")
            If (d Is Nothing) Then
                d = New Directive("TinyPG")
                Me.Directives.Insert(0, d)
            End If
            If Not d.ContainsKey("Namespace") Then
                d.Item("Namespace") = "TinyPG"
            End If
            If Not d.ContainsKey("OutputPath") Then
                d.Item("OutputPath") = "./"
            End If
            If Not d.ContainsKey("Language") Then
                d.Item("OutputPath") = "C#"
            End If
            If Not d.ContainsKey("Language") Then
                d.Item("Language") = "C#"
            End If
            If Not d.ContainsKey("TemplatePath") Then
                Dim lngName As String = d.Item("Language").ToLower(CultureInfo.InvariantCulture)
                If ((Not lngName Is Nothing) AndAlso (((lngName = "visualbasic") OrElse (lngName = "vbnet")) OrElse ((lngName = "vb.net") OrElse (lngName = "vb")))) Then
                    d.Item("TemplatePath") = (AppDomain.CurrentDomain.BaseDirectory & "Templates\VB\")
                Else
                    d.Item("TemplatePath") = (AppDomain.CurrentDomain.BaseDirectory & "Templates\C#\")
                End If
            End If
            d = Me.Directives.Find("Parser")
            If (d Is Nothing) Then
                d = New Directive("Parser")
                Me.Directives.Insert(1, d)
            End If
            If Not d.ContainsKey("Generate") Then
                d.Item("Generate") = "True"
            End If
            d = Me.Directives.Find("Scanner")
            If (d Is Nothing) Then
                d = New Directive("Scanner")
                Me.Directives.Insert(1, d)
            End If
            If Not d.ContainsKey("Generate") Then
                d.Item("Generate") = "True"
            End If
            d = Me.Directives.Find("ParseTree")
            If (d Is Nothing) Then
                d = New Directive("ParseTree")
                Me.Directives.Add(d)
            End If
            If Not d.ContainsKey("Generate") Then
                d.Item("Generate") = "True"
            End If
            d = Me.Directives.Find("TextHighlighter")
            If (d Is Nothing) Then
                d = New Directive("TextHighlighter")
                Me.Directives.Add(d)
            End If
            If Not d.ContainsKey("Generate") Then
                d.Item("Generate") = "False"
            End If
        End Sub

    End Class
End Namespace

