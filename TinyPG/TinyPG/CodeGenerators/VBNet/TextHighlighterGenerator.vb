Imports System
Imports System.IO
Imports System.Text
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler

Namespace TinyPG.CodeGenerators.VBNet
    Public Class TextHighlighterGenerator
        Implements ICodeGenerator
        ' Methods
        Friend Sub New()
        End Sub

        Public Function Generate(ByVal Grammar As Grammar, ByVal Debug As Boolean) As String Implements ICodeGenerator.Generate
            If String.IsNullOrEmpty(Grammar.GetTemplatePath) Then
                Return Nothing
            End If
            Dim generatedtext As String = File.ReadAllText((Grammar.GetTemplatePath & Me.FileName))
            Dim tokens As New StringBuilder
            Dim colors As New StringBuilder
            Dim colorindex As Integer = 1
            Dim t As TerminalSymbol
            For Each t In Grammar.GetTerminals
                If Not t.Attributes.ContainsKey("Color") Then
                    Continue For
                End If
                tokens.AppendLine((Helper.Indent(5) & "Case TokenType." & t.Name & ":"))
                tokens.AppendLine(String.Concat(New Object() {Helper.Indent(6), "sb.Append(""{{\cf", colorindex, " "")"}))
                tokens.AppendLine((Helper.Indent(6) & "Exit Select"))
                Dim red As Integer = 0
                Dim green As Integer = 0
                Dim blue As Integer = 0
                Select Case t.Attributes.Item("Color").Length
                    Case 1
                        If TypeOf t.Attributes.Item("Color")(0) Is Long Then
                            Dim v As Integer = Convert.ToInt32(t.Attributes.Item("Color")(0))
                            red = ((v >> &H10) And &HFF)
                            green = ((v >> 8) And &HFF)
                            blue = (v And &HFF)
                        End If
                        Exit Select
                    Case 3
                        If (TypeOf t.Attributes.Item("Color")(0) Is Integer OrElse TypeOf t.Attributes.Item("Color")(0) Is Long) Then
                            red = (Convert.ToInt32(t.Attributes.Item("Color")(0)) And &HFF)
                        End If
                        If (TypeOf t.Attributes.Item("Color")(1) Is Integer OrElse TypeOf t.Attributes.Item("Color")(1) Is Long) Then
                            green = (Convert.ToInt32(t.Attributes.Item("Color")(1)) And &HFF)
                        End If
                        If (TypeOf t.Attributes.Item("Color")(2) Is Integer OrElse TypeOf t.Attributes.Item("Color")(2) Is Long) Then
                            blue = (Convert.ToInt32(t.Attributes.Item("Color")(2)) And &HFF)
                        End If
                        Exit Select
                End Select
                colors.Append(String.Format("\red{0}\green{1}\blue{2};", red, green, blue))
                colorindex += 1
            Next
            generatedtext = generatedtext.Replace("<%HightlightTokens%>", tokens.ToString).Replace("<%RtfColorPalette%>", colors.ToString)
            If Debug Then
                generatedtext = generatedtext.Replace("<%Namespace%>", "TinyPG.Debug")
            Else
                generatedtext = generatedtext.Replace("<%Namespace%>", Grammar.Directives.Item("TinyPG").Item("Namespace"))
            End If
            Return generatedtext
        End Function


        ' Properties
        Public ReadOnly Property FileName As String Implements ICodeGenerator.FileName
            Get
                Return "TextHighlighter.vb"
            End Get
        End Property

    End Class
End Namespace

