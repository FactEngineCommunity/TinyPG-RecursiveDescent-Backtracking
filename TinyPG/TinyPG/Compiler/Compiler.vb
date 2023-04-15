Imports System
Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports TinyPG.CodeGenerators
Imports TinyPG.Debug

Namespace TinyPG.Compiler
    Public Class Compiler

        Private errorsFiled As List(Of String)
        Private isCompliedField As Boolean
        Private isParsedField As Boolean
        Private parserCodeField As String
        Private [assembly] As Assembly
        Private Grammar As Grammar
        Private parseTreeCodeField As String
        Private scannerCodeField As String
        Public ResultTree As Object

        ' Methods
        Public Sub New()
            Me.IsCompiled = False
            Me.Errors = New List(Of String)
        End Sub

        Private Sub BuildCode()
            Dim language As String = Me.Grammar.Directives.Item("TinyPG").Item("Language")
            Dim provider As CodeDomProvider = CodeGeneratorFactory.CreateCodeDomProvider(language)
            Dim compilerparams As New CompilerParameters With { _
                .GenerateInMemory = True, _
                .GenerateExecutable = False _
            }
            compilerparams.ReferencedAssemblies.Add("System.dll")
            compilerparams.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            compilerparams.ReferencedAssemblies.Add("System.Drawing.dll")
            compilerparams.ReferencedAssemblies.Add("System.Xml.dll")
            Dim tinypgfile As String = Assembly.GetExecutingAssembly.Location
            compilerparams.ReferencedAssemblies.Add(tinypgfile)
            Dim sources As New List(Of String)
            Dim d As Directive
            Dim liInd As Integer = 0
            For Each d In Me.Grammar.Directives
                Dim generator As TinyPG.CodeGenerators.ICodeGenerator = CodeGeneratorFactory.CreateGenerator(d.Name, language)
                If ((Not generator Is Nothing) AndAlso (d.Item("Generate").ToLower = "true")) Then
                    sources.Add(generator.Generate(Me.Grammar, True))
                End If
                'liInd += 1  'For debugging. Added by VM
                'If liInd = 3 Then Exit For
            Next
            If (sources.Count > 0) Then
                Dim Result As CompilerResults = provider.CompileAssemblyFromSource(compilerparams, sources.ToArray)
                If (Result.Errors.Count > 0) Then
                    Dim o As CompilerError
                    For Each o In Result.Errors
                        If Not o.IsWarning Then
                            Me.Errors.Add((o.ErrorText & " on line " & o.Line.ToString))
                        End If
                    Next
                Else
                    Me.assembly = Result.CompiledAssembly
                End If
            End If

        End Sub

        Public Sub Compile(ByVal grammar As Grammar)
            Me.IsParsed = False
            Me.IsCompiled = False
            Me.Errors = New List(Of String)
            If (grammar Is Nothing) Then
                Throw New ArgumentNullException("grammar", "Grammar may not be null")
            End If
            Me.Grammar = grammar
            grammar.Preprocess
            Me.IsParsed = True
            Me.BuildCode
            If (Me.Errors.Count = 0) Then
                Me.IsCompiled = True
            End If
        End Sub

        Public Function Run(ByVal input As String) As CompilerResult
            Return Me.Run(input, Nothing)
        End Function

        Public Function Run(ByVal input As String, ByVal textHighlight As RichTextBox) As CompilerResult
            Dim compilerresult As New CompilerResult
            Dim output As String = Nothing
            If (Me.assembly Is Nothing) Then
                Return Nothing
            End If
            Dim scannerinstance As Object = Me.assembly.CreateInstance("TinyPG.Debug.Scanner")
            Dim scanner As Type = scannerinstance.GetType
            Dim parserinstance As Object = DirectCast(Me.assembly.CreateInstance("TinyPG.Debug.Parser", True, BindingFlags.CreateInstance, Nothing, New Object() { scannerinstance }, Nothing, Nothing), IParser)
            Dim treeinstance As Object = parserinstance.GetType.InvokeMember("Parse", BindingFlags.InvokeMethod, Nothing, parserinstance, New Object() { input })
            Dim itree As IParseTree = TryCast(treeinstance, IParseTree)
            Me.ResultTree = treeinstance
            compilerresult.ParseTree = itree
            Dim errors As List(Of IParseError) = DirectCast(treeinstance.GetType.InvokeMember("Errors", BindingFlags.GetField, Nothing, treeinstance, Nothing), List(Of IParseError))
            If errors.Count = 0 Then
                If ((Not textHighlight Is Nothing) AndAlso (errors.Count = 0)) Then
                    Dim highlighterinstance As Object = Me.assembly.CreateInstance("TinyPG.Debug.TextHighlighter", True, BindingFlags.CreateInstance, Nothing, New Object() {textHighlight, scannerinstance, parserinstance}, Nothing, Nothing)
                    If (Not highlighterinstance Is Nothing) Then
                        output = (output & "Highlighting input..." & ChrW(13) & ChrW(10))
                        Dim highlightertype As Type = highlighterinstance.GetType

                        highlightertype.InvokeMember("HighlightText", BindingFlags.InvokeMethod, Nothing, highlighterinstance, Nothing)
                        Thread.Sleep(20)
                        highlightertype.InvokeMember("Dispose", BindingFlags.InvokeMethod, Nothing, highlighterinstance, Nothing)
                    End If
                End If
            End If
            If (errors.Count > 0) Then
                Dim err As IParseError
                For Each err In errors
                    output = (output & err.Message & ChrW(13) & ChrW(10))
                Next
            Else
                output = (output & "Parse was successful." & ChrW(13) & ChrW(10) & "Evaluating...")
                Try
                    compilerresult.Value = itree.Eval(Nothing)
                    output = (output & ChrW(13) & ChrW(10) & "Result: " & If((compilerresult.Value Is Nothing), "null", compilerresult.Value.ToString))
                Catch exc As Exception
                    output = ((output & ChrW(13) & ChrW(10) & "Exception occurred: " & exc.Message) & ChrW(13) & ChrW(10) & "Stacktrace: " & exc.StackTrace)
                End Try
            End If
            compilerresult.Output = output.ToString
            Return compilerresult
        End Function


        ' Properties
        Public Property Errors As List(Of String)
            Get
                Return Me.errorsFiled
            End Get
            Set(ByVal value As List(Of String))
                Me.errorsFiled = value
            End Set
        End Property

        Public Property IsCompiled As Boolean
            Get
                Return Me.isCompliedField
            End Get
            Set(ByVal value As Boolean)
                Me.isCompliedField = value
            End Set
        End Property

        Public Property IsParsed As Boolean
            Get
                Return Me.isParsedField
            End Get
            Set(ByVal value As Boolean)
                Me.isCompliedField = value
            End Set
        End Property

        Public Property ParserCode As String
            Get
                Return Me.parserCodeField
            End Get
            Set(ByVal value As String)
                Me.parserCodeField = value
            End Set

        End Property

        Public Property ParseTreeCode As String
            Get
                Return Me.parseTreeCodeField
            End Get
            Set(ByVal value As String)
                Me.parseTreeCodeField = value
            End Set
        End Property

        Public Property ScannerCode As String
            Get
                Return Me.scannerCodeField
            End Get
            Set(ByVal value As String)
                Me.scannerCodeField = value
            End Set
        End Property

    End Class
End Namespace

