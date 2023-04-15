Imports Microsoft.CSharp
Imports Microsoft.VisualBasic
Imports System
Imports System.CodeDom.Compiler
Imports System.Globalization
Imports TinyPG.CodeGenerators.CSharp
Imports TinyPG.CodeGenerators.VBNet

Namespace TinyPG.CodeGenerators
    Public Class CodeGeneratorFactory
        ' Methods
        Public Shared Function CreateCodeDomProvider(ByVal language As String) As CodeDomProvider
            Dim lngName As String = language.ToLower(CultureInfo.InvariantCulture)
            If ((Not lngName Is Nothing) AndAlso (((lngName = "visualbasic") OrElse (lngName = "vbnet")) OrElse ((lngName = "vb.net") OrElse (lngName = "vb")))) Then
                Return New VBCodeProvider
            End If
            Return New CSharpCodeProvider
        End Function

        Public Shared Function CreateGenerator(ByVal generator As String, ByVal language As String) As ICodeGenerator
            If (CodeGeneratorFactory.GetSupportedLanguage(language) = SupportedLanguage.VBNet) Then
                Select Case generator
                    Case "Parser"
                        Return New TinyPG.CodeGenerators.VBNet.ParserGenerator
                    Case "Scanner"
                        Return New TinyPG.CodeGenerators.VBNet.ScannerGenerator
                    Case "ParseTree"
                        Return New TinyPG.CodeGenerators.VBNet.ParseTreeGenerator
                    Case "TextHighlighter"
                        Return New TinyPG.CodeGenerators.VBNet.TextHighlighterGenerator
                End Select
            Else
                Select Case generator
                    Case "Parser"
                        Return New TinyPG.CodeGenerators.CSharp.ParserGenerator
                    Case "Scanner"
                        Return New TinyPG.CodeGenerators.CSharp.ScannerGenerator
                    Case "ParseTree"
                        Return New TinyPG.CodeGenerators.CSharp.ParseTreeGenerator
                    Case "TextHighlighter"
                        Return New TinyPG.CodeGenerators.CSharp.TextHighlighterGenerator
                End Select
            End If
            Return Nothing
        End Function

        Public Shared Function GetSupportedLanguage(ByVal language As String) As SupportedLanguage
            Dim lngName As String = language.ToLower(CultureInfo.InvariantCulture)
            If ((Not lngName Is Nothing) AndAlso (((lngName = "visualbasic") OrElse (lngName = "vbnet")) OrElse ((lngName = "vb.net") OrElse (lngName = "vb")))) Then
                Return SupportedLanguage.VBNet
            End If
            Return SupportedLanguage.CSharp
        End Function

    End Class
End Namespace

