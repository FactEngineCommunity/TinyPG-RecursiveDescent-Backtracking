Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports System.Xml
Imports Microsoft.VisualBasic
Imports TinyPG.CodeGenerators
Imports TinyPG.Compiler
Imports TinyPG.Controls
Imports TinyPG.Debug
Imports TinyPG.Highlighter
Imports System.Environment

Namespace TinyPG
    Public Class MainForm
        Inherits Form

        Dim mfrmFindDialog As New DlgFind()

        ' Methods
        Public Sub New()
            Me.InitializeComponent()
            Me.IsDirty = False
            Me.compiler = Nothing
            Me.GrammarFile = Nothing
            AddHandler MyBase.Disposed, New EventHandler(AddressOf Me.MainForm_Disposed)
        End Sub

        Private Sub aboutTinyParserGeneratorToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles aboutTinyParserGeneratorToolStripMenuItem.Click
            Me.AboutTinyPG()
        End Sub

        Private Sub AboutTinyPG()
            Dim about As New StringBuilder
            about.AppendLine((AssemblyInfo.ProductName & " v" & Application.ProductVersion))
            about.AppendLine(AssemblyInfo.CopyRightsDetail)
            about.AppendLine()
            about.AppendLine("For more information about the author")
            about.AppendLine("or TinyPG visit www.codeproject.com")
            Me.outputFloaty.Show()
            Me.tabOutput.SelectedIndex = 0
            Me.textOutput.Text = about.ToString
        End Sub

        Private Sub checker_UpdateSyntax(ByVal sender As Object, ByVal e As EventArgs)
            If Not (Not MyBase.InvokeRequired OrElse MyBase.IsDisposed) Then
                MyBase.Invoke(New EventHandler(AddressOf Me.checker_UpdateSyntax), New Object() {sender, e})
            Else
                Me.marker.MarkWords()
                If ((Not Me.checker.Grammar Is Nothing) AndAlso Not Me.codecomplete.Visible) Then
                    SyncLock Me.checker.Grammar
                        Dim startAdded As Boolean = False
                        Me.codecomplete.WordList.Items.Clear()
                        Dim s As Symbol
                        For Each s In Me.checker.Grammar.Symbols
                            Me.codecomplete.WordList.Items.Add(s.Name)
                            If (s.Name = "Start") Then
                                startAdded = True
                            End If
                        Next
                        If Not startAdded Then
                            Me.codecomplete.WordList.Items.Add("Start")
                        End If
                    End SyncLock
                End If
            End If
        End Sub

        Private Sub codeblocksToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles codeblocksToolStripMenuItem.Click
            MainForm.NotepadViewFile((AppDomain.CurrentDomain.BaseDirectory & "Examples\simple expression2.tpg"))
        End Sub

        Private Sub CompileGrammar()
            Dim starttimer As DateTime = DateTime.Now
            If String.IsNullOrEmpty(Me.GrammarFile) Then
                Me.SaveGrammarAs()
            End If
            If Not String.IsNullOrEmpty(Me.GrammarFile) Then
                Me.compiler = New TinyPG.Compiler.Compiler
                Dim output As New StringBuilder
                Me.tvParsetree.Nodes.Clear()
                Dim input As String = Me.textEditor.Text
                Dim scanner As New Scanner
                Dim tree As ParseTree = New Parser(scanner).Parse(input, New GrammarTree)
                If (tree.Errors.Count > 0) Then
                    Dim errorMsg As ParseError
                    For Each errorMsg In tree.Errors
                        output.AppendLine(errorMsg.Message)
                    Next
                    output.AppendLine("Syntax errors in grammar found.")
                    If (tree.Errors.Count > 0) Then
                        Me.textEditor.Select(tree.Errors.Item(0).Position, If((tree.Errors.Item(0).Length > 0), tree.Errors.Item(0).Length, 1))
                    End If
                Else
                    Me.grammar = DirectCast(tree.Eval(New Object(0 - 1) {}), Grammar)
                    Me.grammar.Preprocess()
                    If (tree.Errors.Count = 0) Then
                        output.AppendLine(Me.grammar.PrintGrammar)
                        output.AppendLine(Me.grammar.PrintFirsts)
                        output.AppendLine("Parse successful!" & ChrW(13) & ChrW(10))
                    End If
                End If
                If (Not Me.grammar Is Nothing) Then
                    Me.SetHighlighterLanguage(Me.grammar.Directives.Item("TinyPG").Item("Language"))
                    If (tree.Errors.Count > 0) Then
                        Dim errorMsg As ParseError
                        For Each errorMsg In tree.Errors
                            output.AppendLine(errorMsg.Message)
                        Next
                        output.AppendLine("Semantic errors in grammar found.")
                        If (tree.Errors.Count > 0) Then
                            Me.textEditor.Select(tree.Errors.Item(0).Position, If((tree.Errors.Item(0).Length > 0), tree.Errors.Item(0).Length, 1))
                        End If
                    Else
                        output.AppendLine("Building code...")
                        Me.compiler.Compile(Me.grammar)
                        If Not Me.compiler.IsCompiled Then
                            Dim err As String
                            For Each err In Me.compiler.Errors
                                output.AppendLine(err)
                            Next
                            output.AppendLine("Compilation contains errors, could not compile.")
                        Else
                            Dim span As TimeSpan = DateTime.Now.Subtract(starttimer)
                            output.AppendLine(("Compilation successfull in " & span.TotalMilliseconds & "ms."))
                        End If
                    End If
                End If
                Me.textOutput.Text = output.ToString
                Me.textOutput.Select(Me.textOutput.Text.Length, 0)
                Me.textOutput.ScrollToCaret()
                If ((Not Me.grammar Is Nothing) AndAlso (tree.Errors.Count = 0)) Then
                    Dim language As String = Me.grammar.Directives.Item("TinyPG").Item("Language")
                    Dim d As Directive
                    For Each d In Me.grammar.Directives
                        Dim generator As ICodeGenerator = CodeGeneratorFactory.CreateGenerator(d.Name, language)
                        If ((Not generator Is Nothing) AndAlso (d.Item("Generate").ToLower = "true")) Then
                            Dim code As String = generator.Generate(Me.grammar, False)
                            File.WriteAllText((Me.grammar.GetOutputPath & generator.FileName), generator.Generate(Me.grammar, False))
                        End If
                    Next
                End If
            End If
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If (disposing AndAlso (Not Me.components Is Nothing)) Then
                Me.checker.Dispose()
                Me.marker.Dispose()
                Me.components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub EvaluateExpression()
            Me.textOutput.Text = "Parsing expression..." & ChrW(13) & ChrW(10)
            Try
                If Not ((Not Me.IsDirty AndAlso (Not Me.compiler Is Nothing)) AndAlso Me.compiler.IsCompiled) Then
                    Me.CompileGrammar()
                End If
                If Not String.IsNullOrEmpty(Me.GrammarFile) Then
                    If ((Not Me.compiler Is Nothing) AndAlso (Me.compiler.Errors.Count = 0)) Then
                        Me.SaveGrammar(Me.GrammarFile)
                    End If
                    Dim result As New CompilerResult
                    If Me.compiler.IsCompiled Then
                        result = Me.compiler.Run(Me.textInput.Text, Me.textInput)
                        Me.textOutput.Text = (Me.textOutput.Text & result.Output)
                        ParseTreeViewer.Populate(Me.tvParsetree, result.ParseTree)

                        Me.ListBoxOptionals.Items.Clear()
                        For Each lrParseError As Object In Me.compiler.ResultTree.Optionals
                            Me.ListBoxOptionals.Items.Add(lrParseError.ExpectedToken.ToString)
                        Next
                    End If
                End If
            Catch exc As Exception
                Dim output As String = Me.textOutput.Text
                Me.textOutput.Text = String.Concat(New String() {output, "An exception occured compiling the assembly: " & ChrW(13) & ChrW(10), exc.Message, ChrW(13) & ChrW(10), exc.StackTrace})
            End Try
        End Sub

        Private Sub exitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles exitToolStripMenuItem.Click
            MyBase.Close()
            Application.Exit()
        End Sub

        Private Sub expressionEvaluatorToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles expressionEvaluatorToolStripMenuItem.Click
            Me.inputFloaty.Show()
            Me.textInput.Focus()
        End Sub

        Private Sub expressionEvaluatorToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles expressionEvaluatorToolStripMenuItem1.Click
            MainForm.NotepadViewFile((AppDomain.CurrentDomain.BaseDirectory & "Examples\simple expression1.tpg"))
        End Sub

        Private Sub headerEvaluator_CloseClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.inputFloaty.Hide()
        End Sub

        Private Sub headerOutput_CloseClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.outputFloaty.Hide()
        End Sub

        Private Sub InitializeComponent()
            Me.menuStrip = New System.Windows.Forms.MenuStrip()
            Me.fileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.newToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.openToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
            Me.saveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.saveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
            Me.exitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.FindToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.viewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.regexToolToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.outputToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.parsetreeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.expressionEvaluatorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.toolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.parseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.menuToolsGenerate = New System.Windows.Forms.ToolStripMenuItem()
            Me.toolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
            Me.viewParserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.viewScannerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.viewParseTreeCodeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.helpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.aboutTinyParserGeneratorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.examplesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.expressionEvaluatorToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
            Me.codeblocksToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.theTinyPGGrammarHighlighterV12ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.theTinyPGGrammarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.theTinyPGGrammarV10ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.statusStrip = New System.Windows.Forms.StatusStrip()
            Me.statusLabel = New System.Windows.Forms.ToolStripStatusLabel()
            Me.toolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
            Me.statusLine = New System.Windows.Forms.ToolStripStatusLabel()
            Me.toolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
            Me.statusCol = New System.Windows.Forms.ToolStripStatusLabel()
            Me.toolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
            Me.statusPos = New System.Windows.Forms.ToolStripStatusLabel()
            Me.toolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
            Me.textEditor = New System.Windows.Forms.RichTextBox()
            Me.splitterBottom = New System.Windows.Forms.Splitter()
            Me.splitterRight = New System.Windows.Forms.Splitter()
            Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
            Me.folderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
            Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
            Me.panelOutput = New System.Windows.Forms.Panel()
            Me.tabOutput = New TinyPG.Controls.TabControlEx()
            Me.tabPage1 = New System.Windows.Forms.TabPage()
            Me.textOutput = New System.Windows.Forms.RichTextBox()
            Me.tabPage2 = New System.Windows.Forms.TabPage()
            Me.tvParsetree = New System.Windows.Forms.TreeView()
            Me.tabPage3 = New System.Windows.Forms.TabPage()
            Me.regExControl = New TinyPG.Controls.RegExControl()
            Me.headerOutput = New TinyPG.Controls.HeaderLabel()
            Me.panelInput = New System.Windows.Forms.Panel()
            Me.textInput = New System.Windows.Forms.RichTextBox()
            Me.headerEvaluator = New TinyPG.Controls.HeaderLabel()
            Me.Optionals = New System.Windows.Forms.TabPage()
            Me.ListBoxOptionals = New System.Windows.Forms.ListBox()
            Me.menuStrip.SuspendLayout()
            Me.statusStrip.SuspendLayout()
            Me.panelOutput.SuspendLayout()
            Me.tabOutput.SuspendLayout()
            Me.tabPage1.SuspendLayout()
            Me.tabPage2.SuspendLayout()
            Me.tabPage3.SuspendLayout()
            Me.panelInput.SuspendLayout()
            Me.Optionals.SuspendLayout()
            Me.SuspendLayout()
            '
            'menuStrip
            '
            Me.menuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.fileToolStripMenuItem, Me.EditToolStripMenuItem, Me.viewToolStripMenuItem, Me.toolsToolStripMenuItem, Me.helpToolStripMenuItem})
            Me.menuStrip.Location = New System.Drawing.Point(0, 0)
            Me.menuStrip.Name = "menuStrip"
            Me.menuStrip.Size = New System.Drawing.Size(1037, 24)
            Me.menuStrip.TabIndex = 0
            Me.menuStrip.Text = "menuStrip"
            '
            'fileToolStripMenuItem
            '
            Me.fileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.newToolStripMenuItem, Me.openToolStripMenuItem, Me.toolStripSeparator1, Me.saveToolStripMenuItem, Me.saveAsToolStripMenuItem, Me.toolStripSeparator2, Me.exitToolStripMenuItem})
            Me.fileToolStripMenuItem.Name = "fileToolStripMenuItem"
            Me.fileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
            Me.fileToolStripMenuItem.Text = "&File"
            '
            'newToolStripMenuItem
            '
            Me.newToolStripMenuItem.Name = "newToolStripMenuItem"
            Me.newToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
            Me.newToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
            Me.newToolStripMenuItem.Text = "&New"
            '
            'openToolStripMenuItem
            '
            Me.openToolStripMenuItem.Name = "openToolStripMenuItem"
            Me.openToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
            Me.openToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
            Me.openToolStripMenuItem.Text = "&Open..."
            '
            'toolStripSeparator1
            '
            Me.toolStripSeparator1.Name = "toolStripSeparator1"
            Me.toolStripSeparator1.Size = New System.Drawing.Size(152, 6)
            '
            'saveToolStripMenuItem
            '
            Me.saveToolStripMenuItem.Name = "saveToolStripMenuItem"
            Me.saveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
            Me.saveToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
            Me.saveToolStripMenuItem.Text = "&Save"
            '
            'saveAsToolStripMenuItem
            '
            Me.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem"
            Me.saveAsToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
            Me.saveAsToolStripMenuItem.Text = "Save &As..."
            '
            'toolStripSeparator2
            '
            Me.toolStripSeparator2.Name = "toolStripSeparator2"
            Me.toolStripSeparator2.Size = New System.Drawing.Size(152, 6)
            '
            'exitToolStripMenuItem
            '
            Me.exitToolStripMenuItem.Name = "exitToolStripMenuItem"
            Me.exitToolStripMenuItem.Size = New System.Drawing.Size(155, 22)
            Me.exitToolStripMenuItem.Text = "E&xit"
            '
            'EditToolStripMenuItem
            '
            Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FindToolStripMenuItem})
            Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
            Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
            Me.EditToolStripMenuItem.Text = "&Edit"
            '
            'FindToolStripMenuItem
            '
            Me.FindToolStripMenuItem.Name = "FindToolStripMenuItem"
            Me.FindToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
            Me.FindToolStripMenuItem.Size = New System.Drawing.Size(137, 22)
            Me.FindToolStripMenuItem.Text = "&Find"
            '
            'viewToolStripMenuItem
            '
            Me.viewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.regexToolToolStripMenuItem, Me.outputToolStripMenuItem, Me.parsetreeToolStripMenuItem, Me.expressionEvaluatorToolStripMenuItem})
            Me.viewToolStripMenuItem.Name = "viewToolStripMenuItem"
            Me.viewToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
            Me.viewToolStripMenuItem.Text = "&View"
            '
            'regexToolToolStripMenuItem
            '
            Me.regexToolToolStripMenuItem.Name = "regexToolToolStripMenuItem"
            Me.regexToolToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
            Me.regexToolToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.regexToolToolStripMenuItem.Text = "Regex tool"
            '
            'outputToolStripMenuItem
            '
            Me.outputToolStripMenuItem.Name = "outputToolStripMenuItem"
            Me.outputToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
            Me.outputToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.outputToolStripMenuItem.Text = "&Output"
            '
            'parsetreeToolStripMenuItem
            '
            Me.parsetreeToolStripMenuItem.Name = "parsetreeToolStripMenuItem"
            Me.parsetreeToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.T), System.Windows.Forms.Keys)
            Me.parsetreeToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.parsetreeToolStripMenuItem.Text = "Parse &tree"
            '
            'expressionEvaluatorToolStripMenuItem
            '
            Me.expressionEvaluatorToolStripMenuItem.Name = "expressionEvaluatorToolStripMenuItem"
            Me.expressionEvaluatorToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
            Me.expressionEvaluatorToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.expressionEvaluatorToolStripMenuItem.Text = "&Expression evaluator"
            '
            'toolsToolStripMenuItem
            '
            Me.toolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.parseToolStripMenuItem, Me.menuToolsGenerate, Me.toolStripMenuItem1, Me.viewParserToolStripMenuItem, Me.viewScannerToolStripMenuItem, Me.viewParseTreeCodeToolStripMenuItem})
            Me.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem"
            Me.toolsToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
            Me.toolsToolStripMenuItem.Text = "&Build"
            '
            'parseToolStripMenuItem
            '
            Me.parseToolStripMenuItem.Name = "parseToolStripMenuItem"
            Me.parseToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5
            Me.parseToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
            Me.parseToolStripMenuItem.Text = "Generate && &Run"
            '
            'menuToolsGenerate
            '
            Me.menuToolsGenerate.Name = "menuToolsGenerate"
            Me.menuToolsGenerate.ShortcutKeys = System.Windows.Forms.Keys.F6
            Me.menuToolsGenerate.Size = New System.Drawing.Size(180, 22)
            Me.menuToolsGenerate.Text = "&Generate"
            '
            'toolStripMenuItem1
            '
            Me.toolStripMenuItem1.Name = "toolStripMenuItem1"
            Me.toolStripMenuItem1.Size = New System.Drawing.Size(177, 6)
            '
            'viewParserToolStripMenuItem
            '
            Me.viewParserToolStripMenuItem.Name = "viewParserToolStripMenuItem"
            Me.viewParserToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
            Me.viewParserToolStripMenuItem.Text = "View &Parser code"
            '
            'viewScannerToolStripMenuItem
            '
            Me.viewScannerToolStripMenuItem.Name = "viewScannerToolStripMenuItem"
            Me.viewScannerToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
            Me.viewScannerToolStripMenuItem.Text = "View &Scanner code"
            '
            'viewParseTreeCodeToolStripMenuItem
            '
            Me.viewParseTreeCodeToolStripMenuItem.Name = "viewParseTreeCodeToolStripMenuItem"
            Me.viewParseTreeCodeToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
            Me.viewParseTreeCodeToolStripMenuItem.Text = "View Parse&Tree code"
            '
            'helpToolStripMenuItem
            '
            Me.helpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.aboutTinyParserGeneratorToolStripMenuItem, Me.examplesToolStripMenuItem})
            Me.helpToolStripMenuItem.Name = "helpToolStripMenuItem"
            Me.helpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
            Me.helpToolStripMenuItem.Text = "&Help"
            '
            'aboutTinyParserGeneratorToolStripMenuItem
            '
            Me.aboutTinyParserGeneratorToolStripMenuItem.Name = "aboutTinyParserGeneratorToolStripMenuItem"
            Me.aboutTinyParserGeneratorToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.aboutTinyParserGeneratorToolStripMenuItem.Text = "&About Tiny Parser Generator"
            '
            'examplesToolStripMenuItem
            '
            Me.examplesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.expressionEvaluatorToolStripMenuItem1, Me.codeblocksToolStripMenuItem, Me.theTinyPGGrammarHighlighterV12ToolStripMenuItem, Me.theTinyPGGrammarToolStripMenuItem, Me.theTinyPGGrammarV10ToolStripMenuItem})
            Me.examplesToolStripMenuItem.Name = "examplesToolStripMenuItem"
            Me.examplesToolStripMenuItem.Size = New System.Drawing.Size(222, 22)
            Me.examplesToolStripMenuItem.Text = "&Examples"
            '
            'expressionEvaluatorToolStripMenuItem1
            '
            Me.expressionEvaluatorToolStripMenuItem1.Name = "expressionEvaluatorToolStripMenuItem1"
            Me.expressionEvaluatorToolStripMenuItem1.Size = New System.Drawing.Size(273, 22)
            Me.expressionEvaluatorToolStripMenuItem1.Text = "Simple Expression evaluator"
            '
            'codeblocksToolStripMenuItem
            '
            Me.codeblocksToolStripMenuItem.Name = "codeblocksToolStripMenuItem"
            Me.codeblocksToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
            Me.codeblocksToolStripMenuItem.Text = "Simple Expression calculator"
            '
            'theTinyPGGrammarHighlighterV12ToolStripMenuItem
            '
            Me.theTinyPGGrammarHighlighterV12ToolStripMenuItem.Name = "theTinyPGGrammarHighlighterV12ToolStripMenuItem"
            Me.theTinyPGGrammarHighlighterV12ToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
            Me.theTinyPGGrammarHighlighterV12ToolStripMenuItem.Text = "The TinyPG Grammar Highlighter v1.2"
            '
            'theTinyPGGrammarToolStripMenuItem
            '
            Me.theTinyPGGrammarToolStripMenuItem.Name = "theTinyPGGrammarToolStripMenuItem"
            Me.theTinyPGGrammarToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
            Me.theTinyPGGrammarToolStripMenuItem.Text = "The TinyPG Grammar v1.1"
            '
            'theTinyPGGrammarV10ToolStripMenuItem
            '
            Me.theTinyPGGrammarV10ToolStripMenuItem.Name = "theTinyPGGrammarV10ToolStripMenuItem"
            Me.theTinyPGGrammarV10ToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
            Me.theTinyPGGrammarV10ToolStripMenuItem.Text = "The TinyPG Grammar v1.0"
            '
            'statusStrip
            '
            Me.statusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusLabel, Me.toolStripStatusLabel1, Me.statusLine, Me.toolStripStatusLabel2, Me.statusCol, Me.toolStripStatusLabel4, Me.statusPos, Me.toolStripStatusLabel3})
            Me.statusStrip.Location = New System.Drawing.Point(0, 624)
            Me.statusStrip.Name = "statusStrip"
            Me.statusStrip.Size = New System.Drawing.Size(1037, 22)
            Me.statusStrip.TabIndex = 1
            Me.statusStrip.Text = "statusStrip1"
            '
            'statusLabel
            '
            Me.statusLabel.Name = "statusLabel"
            Me.statusLabel.Size = New System.Drawing.Size(751, 17)
            Me.statusLabel.Spring = True
            Me.statusLabel.Text = "Ready"
            Me.statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'toolStripStatusLabel1
            '
            Me.toolStripStatusLabel1.Name = "toolStripStatusLabel1"
            Me.toolStripStatusLabel1.Size = New System.Drawing.Size(20, 17)
            Me.toolStripStatusLabel1.Text = "Ln"
            Me.toolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'statusLine
            '
            Me.statusLine.AutoSize = False
            Me.statusLine.Name = "statusLine"
            Me.statusLine.Size = New System.Drawing.Size(50, 17)
            Me.statusLine.Text = "-"
            Me.statusLine.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'toolStripStatusLabel2
            '
            Me.toolStripStatusLabel2.Name = "toolStripStatusLabel2"
            Me.toolStripStatusLabel2.Size = New System.Drawing.Size(25, 17)
            Me.toolStripStatusLabel2.Text = "Col"
            Me.toolStripStatusLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
            '
            'statusCol
            '
            Me.statusCol.AutoSize = False
            Me.statusCol.Name = "statusCol"
            Me.statusCol.Size = New System.Drawing.Size(50, 17)
            Me.statusCol.Text = "-"
            Me.statusCol.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'toolStripStatusLabel4
            '
            Me.toolStripStatusLabel4.Name = "toolStripStatusLabel4"
            Me.toolStripStatusLabel4.Size = New System.Drawing.Size(26, 17)
            Me.toolStripStatusLabel4.Text = "Pos"
            '
            'statusPos
            '
            Me.statusPos.AutoSize = False
            Me.statusPos.Name = "statusPos"
            Me.statusPos.Size = New System.Drawing.Size(50, 17)
            Me.statusPos.Text = "-"
            Me.statusPos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'toolStripStatusLabel3
            '
            Me.toolStripStatusLabel3.AutoSize = False
            Me.toolStripStatusLabel3.Name = "toolStripStatusLabel3"
            Me.toolStripStatusLabel3.Size = New System.Drawing.Size(50, 17)
            Me.toolStripStatusLabel3.Text = "INS"
            '
            'textEditor
            '
            Me.textEditor.AcceptsTab = True
            Me.textEditor.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.textEditor.Dock = System.Windows.Forms.DockStyle.Fill
            Me.textEditor.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.textEditor.HideSelection = False
            Me.textEditor.Location = New System.Drawing.Point(0, 24)
            Me.textEditor.Name = "textEditor"
            Me.textEditor.Size = New System.Drawing.Size(712, 417)
            Me.textEditor.TabIndex = 3
            Me.textEditor.Text = ""
            Me.textEditor.WordWrap = False
            '
            'splitterBottom
            '
            Me.splitterBottom.BackColor = System.Drawing.SystemColors.InactiveCaption
            Me.splitterBottom.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.splitterBottom.Location = New System.Drawing.Point(0, 441)
            Me.splitterBottom.Name = "splitterBottom"
            Me.splitterBottom.Size = New System.Drawing.Size(712, 5)
            Me.splitterBottom.TabIndex = 5
            Me.splitterBottom.TabStop = False
            '
            'splitterRight
            '
            Me.splitterRight.BackColor = System.Drawing.SystemColors.InactiveCaption
            Me.splitterRight.Dock = System.Windows.Forms.DockStyle.Right
            Me.splitterRight.Location = New System.Drawing.Point(712, 24)
            Me.splitterRight.Name = "splitterRight"
            Me.splitterRight.Size = New System.Drawing.Size(5, 600)
            Me.splitterRight.TabIndex = 7
            Me.splitterRight.TabStop = False
            '
            'openFileDialog
            '
            Me.openFileDialog.DefaultExt = "tpg"
            Me.openFileDialog.Filter = "Grammar files|*.tpg|All files|*.*"
            Me.openFileDialog.Title = "Open Grammar File"
            '
            'folderBrowserDialog
            '
            Me.folderBrowserDialog.RootFolder = System.Environment.SpecialFolder.Favorites
            '
            'saveFileDialog
            '
            Me.saveFileDialog.Filter = "Grammar files|*.tpg|All files|*.*"
            Me.saveFileDialog.Title = "Save Grammar File As"
            '
            'panelOutput
            '
            Me.panelOutput.Controls.Add(Me.tabOutput)
            Me.panelOutput.Controls.Add(Me.headerOutput)
            Me.panelOutput.Dock = System.Windows.Forms.DockStyle.Right
            Me.panelOutput.Location = New System.Drawing.Point(717, 24)
            Me.panelOutput.Name = "panelOutput"
            Me.panelOutput.Size = New System.Drawing.Size(320, 600)
            Me.panelOutput.TabIndex = 8
            '
            'tabOutput
            '
            Me.tabOutput.Alignment = System.Windows.Forms.TabAlignment.Bottom
            Me.tabOutput.Controls.Add(Me.tabPage1)
            Me.tabOutput.Controls.Add(Me.tabPage2)
            Me.tabOutput.Controls.Add(Me.tabPage3)
            Me.tabOutput.Controls.Add(Me.Optionals)
            Me.tabOutput.Dock = System.Windows.Forms.DockStyle.Fill
            Me.tabOutput.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
            Me.tabOutput.Location = New System.Drawing.Point(0, 20)
            Me.tabOutput.Name = "tabOutput"
            Me.tabOutput.Padding = New System.Drawing.Point(10, 3)
            Me.tabOutput.SelectedIndex = 0
            Me.tabOutput.Size = New System.Drawing.Size(320, 580)
            Me.tabOutput.TabIndex = 6
            '
            'tabPage1
            '
            Me.tabPage1.Controls.Add(Me.textOutput)
            Me.tabPage1.Location = New System.Drawing.Point(4, 4)
            Me.tabPage1.Name = "tabPage1"
            Me.tabPage1.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage1.Size = New System.Drawing.Size(312, 554)
            Me.tabPage1.TabIndex = 0
            Me.tabPage1.Text = "Output"
            Me.tabPage1.UseVisualStyleBackColor = True
            '
            'textOutput
            '
            Me.textOutput.BackColor = System.Drawing.SystemColors.Window
            Me.textOutput.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.textOutput.Dock = System.Windows.Forms.DockStyle.Fill
            Me.textOutput.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.textOutput.Location = New System.Drawing.Point(3, 3)
            Me.textOutput.Name = "textOutput"
            Me.textOutput.ReadOnly = True
            Me.textOutput.Size = New System.Drawing.Size(306, 548)
            Me.textOutput.TabIndex = 6
            Me.textOutput.Text = ""
            Me.textOutput.WordWrap = False
            '
            'tabPage2
            '
            Me.tabPage2.Controls.Add(Me.tvParsetree)
            Me.tabPage2.Location = New System.Drawing.Point(4, 4)
            Me.tabPage2.Name = "tabPage2"
            Me.tabPage2.Padding = New System.Windows.Forms.Padding(3)
            Me.tabPage2.Size = New System.Drawing.Size(312, 554)
            Me.tabPage2.TabIndex = 1
            Me.tabPage2.Text = "Parse tree"
            Me.tabPage2.UseVisualStyleBackColor = True
            '
            'tvParsetree
            '
            Me.tvParsetree.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.tvParsetree.Dock = System.Windows.Forms.DockStyle.Fill
            Me.tvParsetree.Location = New System.Drawing.Point(3, 3)
            Me.tvParsetree.Name = "tvParsetree"
            Me.tvParsetree.Size = New System.Drawing.Size(306, 548)
            Me.tvParsetree.TabIndex = 0
            '
            'tabPage3
            '
            Me.tabPage3.Controls.Add(Me.regExControl)
            Me.tabPage3.Location = New System.Drawing.Point(4, 4)
            Me.tabPage3.Name = "tabPage3"
            Me.tabPage3.Size = New System.Drawing.Size(312, 554)
            Me.tabPage3.TabIndex = 2
            Me.tabPage3.Text = "Regex tool"
            Me.tabPage3.UseVisualStyleBackColor = True
            '
            'regExControl
            '
            Me.regExControl.BackColor = System.Drawing.SystemColors.Control
            Me.regExControl.Dock = System.Windows.Forms.DockStyle.Fill
            Me.regExControl.Location = New System.Drawing.Point(0, 0)
            Me.regExControl.Name = "regExControl"
            Me.regExControl.Size = New System.Drawing.Size(312, 554)
            Me.regExControl.TabIndex = 12
            '
            'headerOutput
            '
            Me.headerOutput.Dock = System.Windows.Forms.DockStyle.Top
            Me.headerOutput.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.headerOutput.ForeColor = System.Drawing.SystemColors.GrayText
            Me.headerOutput.Location = New System.Drawing.Point(0, 0)
            Me.headerOutput.Name = "headerOutput"
            Me.headerOutput.Size = New System.Drawing.Size(320, 20)
            Me.headerOutput.TabIndex = 7
            Me.headerOutput.Text = "Output"
            Me.headerOutput.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'panelInput
            '
            Me.panelInput.Controls.Add(Me.textInput)
            Me.panelInput.Controls.Add(Me.headerEvaluator)
            Me.panelInput.Dock = System.Windows.Forms.DockStyle.Bottom
            Me.panelInput.Location = New System.Drawing.Point(0, 446)
            Me.panelInput.Margin = New System.Windows.Forms.Padding(0)
            Me.panelInput.Name = "panelInput"
            Me.panelInput.Size = New System.Drawing.Size(712, 178)
            Me.panelInput.TabIndex = 9
            '
            'textInput
            '
            Me.textInput.AcceptsTab = True
            Me.textInput.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.textInput.DetectUrls = False
            Me.textInput.Dock = System.Windows.Forms.DockStyle.Fill
            Me.textInput.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.textInput.HideSelection = False
            Me.textInput.Location = New System.Drawing.Point(0, 20)
            Me.textInput.Name = "textInput"
            Me.textInput.Size = New System.Drawing.Size(712, 158)
            Me.textInput.TabIndex = 2
            Me.textInput.Text = ""
            Me.textInput.WordWrap = False
            '
            'headerEvaluator
            '
            Me.headerEvaluator.Dock = System.Windows.Forms.DockStyle.Top
            Me.headerEvaluator.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.headerEvaluator.ForeColor = System.Drawing.SystemColors.GrayText
            Me.headerEvaluator.Location = New System.Drawing.Point(0, 0)
            Me.headerEvaluator.Name = "headerEvaluator"
            Me.headerEvaluator.Size = New System.Drawing.Size(712, 20)
            Me.headerEvaluator.TabIndex = 3
            Me.headerEvaluator.Text = "Expression Evaluator"
            Me.headerEvaluator.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'Optionals
            '
            Me.Optionals.Controls.Add(Me.ListBoxOptionals)
            Me.Optionals.Location = New System.Drawing.Point(4, 4)
            Me.Optionals.Name = "Optionals"
            Me.Optionals.Size = New System.Drawing.Size(312, 554)
            Me.Optionals.TabIndex = 3
            Me.Optionals.Text = "Optionals"
            Me.Optionals.UseVisualStyleBackColor = True
            '
            'ListBoxOptionals
            '
            Me.ListBoxOptionals.Dock = System.Windows.Forms.DockStyle.Fill
            Me.ListBoxOptionals.FormattingEnabled = True
            Me.ListBoxOptionals.Location = New System.Drawing.Point(0, 0)
            Me.ListBoxOptionals.Name = "ListBoxOptionals"
            Me.ListBoxOptionals.Size = New System.Drawing.Size(312, 554)
            Me.ListBoxOptionals.TabIndex = 0
            '
            'MainForm
            '
            Me.ClientSize = New System.Drawing.Size(1037, 646)
            Me.Controls.Add(Me.textEditor)
            Me.Controls.Add(Me.splitterBottom)
            Me.Controls.Add(Me.panelInput)
            Me.Controls.Add(Me.splitterRight)
            Me.Controls.Add(Me.panelOutput)
            Me.Controls.Add(Me.menuStrip)
            Me.Controls.Add(Me.statusStrip)
            Me.MainMenuStrip = Me.menuStrip
            Me.Name = "MainForm"
            Me.Text = "Tiny Parser Generator .Net"
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
            Me.menuStrip.ResumeLayout(False)
            Me.menuStrip.PerformLayout()
            Me.statusStrip.ResumeLayout(False)
            Me.statusStrip.PerformLayout()
            Me.panelOutput.ResumeLayout(False)
            Me.tabOutput.ResumeLayout(False)
            Me.tabPage1.ResumeLayout(False)
            Me.tabPage2.ResumeLayout(False)
            Me.tabPage3.ResumeLayout(False)
            Me.panelInput.ResumeLayout(False)
            Me.Optionals.ResumeLayout(False)
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Private Sub inputFloaty_Docking(ByVal sender As Object, ByVal e As EventArgs)
            Me.textEditor.BringToFront()
        End Sub

        Private Sub LoadConfig()
            Try
                Dim configfile As String = (AppDomain.CurrentDomain.BaseDirectory & "TinyPG.config")
                If File.Exists(configfile) Then
                    Dim doc As New XmlDocument
                    doc.Load(configfile)
                    Me.openFileDialog.InitialDirectory = doc.Item("AppSettings").Item("OpenFilePath").Attributes.ItemOf(0).Value
                    Me.saveFileDialog.InitialDirectory = doc.Item("AppSettings").Item("SaveFilePath").Attributes.ItemOf(0).Value
                    Me.GrammarFile = doc.Item("AppSettings").Item("GrammarFile").Attributes.ItemOf(0).Value
                End If
            Catch exception1 As Exception
            End Try
        End Sub

        Private Sub LoadGrammarFile()
            If (Not Me.GrammarFile Is Nothing) Then
                If Not File.Exists(Me.GrammarFile) Then
                    Me.GrammarFile = Nothing
                Else
                    Directory.SetCurrentDirectory(New FileInfo(Me.GrammarFile).DirectoryName)
                    Me.textEditor.Text = File.ReadAllText(Me.GrammarFile)
                    Me.textEditor.ClearUndo()
                    Me.CompileGrammar()
                    Me.textOutput.Text = ""
                    Me.textEditor.Focus()
                    Me.SetStatusbar()
                    Me.textHighlighter.ClearUndo()
                    Me.IsDirty = False
                    Me.SetFormCaption()
                    Me.textEditor.Select(0, 0)
                    Me.checker.Check(Me.textEditor.Text)
                End If
            End If
        End Sub

        Private Sub MainForm_Disposed(ByVal sender As Object, ByVal e As EventArgs)
            RemoveHandler Me.checker.UpdateSyntax, Me.syntaxUpdateChecker
            Me.checker.Dispose()
            Me.marker.Dispose()
        End Sub

        Private Sub MainForm_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
            Me.headerEvaluator.Activate(Me.textInput)
            AddHandler Me.headerEvaluator.CloseClick, New EventHandler(AddressOf Me.headerEvaluator_CloseClick)
            Me.headerOutput.Activate(Me.tabOutput)
            AddHandler Me.headerOutput.CloseClick, New EventHandler(AddressOf Me.headerOutput_CloseClick)
            Me.DockExtender = New DockExtender(Me)
            Me.inputFloaty = Me.DockExtender.Attach(Me.panelInput, Me.headerEvaluator, Me.splitterBottom)
            Me.outputFloaty = Me.DockExtender.Attach(Me.panelOutput, Me.headerOutput, Me.splitterRight)
            AddHandler Me.inputFloaty.Docking, New EventHandler(AddressOf Me.inputFloaty_Docking)
            AddHandler Me.outputFloaty.Docking, New EventHandler(AddressOf Me.inputFloaty_Docking)
            Me.inputFloaty.Hide()
            Me.outputFloaty.Hide()
            Me.textOutput.Text = (AssemblyInfo.ProductName & " v" & Application.ProductVersion & ChrW(13) & ChrW(10))
            Me.textOutput.Text = (Me.textOutput.Text & AssemblyInfo.CopyRightsDetail & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10))
            Me.marker = New TextMarker(Me.textEditor)
            Me.checker = New SyntaxChecker(Me.marker)
            Me.syntaxUpdateChecker = New EventHandler(AddressOf Me.checker_UpdateSyntax)
            AddHandler Me.checker.UpdateSyntax, Me.syntaxUpdateChecker
            Dim thread As Thread = New Thread(New ThreadStart(AddressOf Me.checker.Start))
            thread.Start()
            Me.TextChangedTimer = New System.Windows.Forms.Timer
            AddHandler Me.TextChangedTimer.Tick, New EventHandler(AddressOf Me.TextChangedTimer_Tick)
            Me.codecomplete = New AutoComplete(Me.textEditor)
            Me.codecomplete.Enabled = False
            Me.directivecomplete = New AutoComplete(Me.textEditor)
            Me.directivecomplete.Enabled = False
            Me.directivecomplete.WordList.Items.Add("@ParseTree")
            Me.directivecomplete.WordList.Items.Add("@Parser")
            Me.directivecomplete.WordList.Items.Add("@Scanner")
            Me.directivecomplete.WordList.Items.Add("@TextHighlighter")
            Me.directivecomplete.WordList.Items.Add("@TinyPG")
            Me.directivecomplete.WordList.Items.Add("Generate")
            Me.directivecomplete.WordList.Items.Add("Language")
            Me.directivecomplete.WordList.Items.Add("Namespace")
            Me.directivecomplete.WordList.Items.Add("OutputPath")
            Me.directivecomplete.WordList.Items.Add("TemplatePath")
            Me.highlighterScanner = New TinyPG.Highlighter.Scanner()
            Me.textHighlighter = New TextHighlighter(Me.textEditor, Me.highlighterScanner, New TinyPG.Highlighter.Parser(Me.highlighterScanner))
            AddHandler Me.textHighlighter.SwitchContext, New ContextSwitchEventHandler(AddressOf Me.TextHighlighter_SwitchContext)
            Me.LoadConfig()
            '-------------------------------------------------------------------------------------------------------
            'If the user double-clicked on a TPG file to open TinyPG then use that file as the current GrammarFile
            '-------------------------------------------------------------------------------------------------------
            If psStartupTPGFile <> "" Then
                Me.GrammarFile = psStartupTPGFile
            End If
            If String.IsNullOrEmpty(Me.GrammarFile) Then
                Me.NewGrammar()
            Else
                Me.LoadGrammarFile()
            End If

            Me.inputFloaty.Show()
            Me.textInput.Focus()
        End Sub

        Private Sub menuToolsGenerate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuToolsGenerate.Click
            Me.outputFloaty.Show()
            Me.tabOutput.SelectedIndex = 0
            Me.CompileGrammar()
            If ((Not Me.compiler Is Nothing) AndAlso (Me.compiler.Errors.Count = 0)) Then
                Me.SaveGrammar(Me.GrammarFile)
            End If
        End Sub

        Private Sub NewGrammar()
            Me.GrammarFile = Nothing
            Me.IsDirty = False
            Dim text As String = (String.Concat(New String() {"//", AssemblyInfo.ProductName, " v", Application.ProductVersion, ChrW(13) & ChrW(10)}) & "//" & AssemblyInfo.CopyRightsDetail & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10))
            Me.textEditor.Text = [text]
            Me.textEditor.ClearUndo()
            Me.textOutput.Text = (AssemblyInfo.ProductName & " v" & Application.ProductVersion & ChrW(13) & ChrW(10))
            Me.textOutput.Text = (Me.textOutput.Text & AssemblyInfo.CopyRightsDetail & ChrW(13) & ChrW(10) & ChrW(13) & ChrW(10))
            Me.SetFormCaption()
            Me.SaveConfig()
            Me.textEditor.Select(Me.textEditor.Text.Length, 0)
            Me.IsDirty = False
            Me.textHighlighter.ClearUndo()
            Me.SetFormCaption()
            Me.SetStatusbar()
        End Sub

        Private Sub newToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles newToolStripMenuItem.Click
            If Me.IsDirty Then
                Me.SaveGrammarAs()
            End If
            Me.NewGrammar()
        End Sub

        Private Shared Sub NotepadViewFile(ByVal filename As String)
            Try
                Process.Start("Notepad.exe", filename)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        Private Function OpenGrammar() As String
            If (Me.openFileDialog.ShowDialog(Me) = DialogResult.OK) Then
                Return Me.openFileDialog.FileName
            End If
            Return Nothing
        End Function

        Private Sub openToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles openToolStripMenuItem.Click
            Dim newgrammarfile As String = Me.OpenGrammar
            If ((Not newgrammarfile Is Nothing) AndAlso ((Not Me.IsDirty OrElse (Me.GrammarFile Is Nothing)) OrElse (MessageBox.Show(Me, "You will lose current changes, continue?", "Lose changes", MessageBoxButtons.OKCancel) <> DialogResult.Cancel))) Then
                Me.GrammarFile = newgrammarfile
                Me.LoadGrammarFile()
                Me.SaveConfig()
            End If
        End Sub

        Private Sub outputToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles outputToolStripMenuItem.Click
            Me.outputFloaty.Show()
            Me.tabOutput.SelectedIndex = 0
        End Sub

        Private Sub parseToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles parseToolStripMenuItem.Click
            Me.inputFloaty.Show()
            Me.outputFloaty.Show()
            If ((Me.tabOutput.SelectedIndex <> 0) AndAlso (Me.tabOutput.SelectedIndex <> 1)) Then
                Me.tabOutput.SelectedIndex = 0
            End If
            Me.EvaluateExpression()
        End Sub

        Private Sub parsetreeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles parsetreeToolStripMenuItem.Click
            Me.outputFloaty.Show()
            Me.tabOutput.SelectedIndex = 1
        End Sub

        Private Sub regexToolToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles regexToolToolStripMenuItem.Click
            Me.outputFloaty.Show()
            Me.tabOutput.SelectedIndex = 2
        End Sub

        Private Sub saveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles saveAsToolStripMenuItem.Click
            Me.SaveGrammarAs()
            Me.SaveConfig()
        End Sub

        Private Sub SaveConfig()
            Dim configfile As String = (AppDomain.CurrentDomain.BaseDirectory & "TinyPG.config")
            Dim doc As New XmlDocument
            Dim settings As XmlNode = doc.CreateElement("AppSettings", "TinyPG")
            doc.AppendChild(settings)
            Dim node As XmlNode = doc.CreateElement("OpenFilePath", "TinyPG")
            settings.AppendChild(node)
            node = doc.CreateElement("SaveFilePath", "TinyPG")
            settings.AppendChild(node)
            node = doc.CreateElement("GrammarFile", "TinyPG")
            settings.AppendChild(node)
            Dim attr As XmlAttribute = doc.CreateAttribute("Value")
            settings.Item("OpenFilePath").Attributes.Append(attr)
            If File.Exists(Me.openFileDialog.FileName) Then
                attr.Value = New FileInfo(Me.openFileDialog.FileName).Directory.FullName
            End If
            attr = doc.CreateAttribute("Value")
            settings.Item("SaveFilePath").Attributes.Append(attr)
            If File.Exists(Me.saveFileDialog.FileName) Then
                attr.Value = New FileInfo(Me.saveFileDialog.FileName).Directory.FullName
            End If
            attr = doc.CreateAttribute("Value")
            attr.Value = Me.GrammarFile
            settings.Item("GrammarFile").Attributes.Append(attr)
            doc.Save(configfile)
        End Sub

        Private Sub SaveGrammar(ByVal filename As String)
            If Not String.IsNullOrEmpty(filename) Then
                Me.GrammarFile = filename
                Directory.SetCurrentDirectory(New FileInfo(Me.GrammarFile).DirectoryName)
                Dim text As String = Me.textEditor.Text.Replace(ChrW(10), ChrW(13) & ChrW(10))
                File.WriteAllText(filename, [text])
                Me.IsDirty = False
                Me.SetFormCaption()
            End If
        End Sub

        Private Sub SaveGrammarAs()
            If (Me.saveFileDialog.ShowDialog(Me) = DialogResult.OK) Then
                Me.SaveGrammar(Me.saveFileDialog.FileName)
            End If
        End Sub

        Private Sub saveToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles saveToolStripMenuItem.Click
            If String.IsNullOrEmpty(Me.GrammarFile) Then
                Me.SaveGrammarAs()
            Else
                Me.SaveGrammar(Me.GrammarFile)
            End If
            Me.SaveConfig()
        End Sub

        Private Sub SetFormCaption()
            Me.Text = "@TinyPG - a Tiny Parser Generator .Net"
            If Not ((Not Me.GrammarFile Is Nothing) AndAlso File.Exists(Me.GrammarFile)) Then
                If Me.IsDirty Then
                    Me.Text = (Me.Text & " *")
                End If
            Else
                Dim name As String = New FileInfo(Me.GrammarFile).Name
                Me.Text = (Me.Text & " [" & name & "]")
                If Me.IsDirty Then
                    Me.Text = (Me.Text & " *")
                End If
            End If
        End Sub

        Private Sub SetHighlighterLanguage(ByVal language As String)
            SyncLock TextHighlighter.treelock
                If (CodeGeneratorFactory.GetSupportedLanguage(language) = SupportedLanguage.VBNet) Then
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_STRING) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_STRING)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_SYMBOL) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_SYMBOL)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_COMMENTBLOCK) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_COMMENTBLOCK)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_COMMENTLINE) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_COMMENTLINE)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_KEYWORD) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_KEYWORD)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_NONKEYWORD) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.VB_NONKEYWORD)
                Else
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_STRING) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_STRING)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_SYMBOL) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_SYMBOL)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_COMMENTBLOCK) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_COMMENTBLOCK)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_COMMENTLINE) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_COMMENTLINE)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_KEYWORD) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_KEYWORD)
                    Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.DOTNET_NONKEYWORD) = Me.highlighterScanner.Patterns.Item(TinyPG.Highlighter.TokenType.CS_NONKEYWORD)
                End If
                Me.textHighlighter.HighlightText()
            End SyncLock
        End Sub

        Private Sub SetStatusbar()
            Dim pos As Integer
            If Me.textEditor.Focused Then
                pos = Me.textEditor.SelectionStart
                Me.statusPos.Text = pos.ToString(CultureInfo.InvariantCulture)
                Dim postion As Integer = (pos - Me.textEditor.GetFirstCharIndexOfCurrentLine)
                Me.statusCol.Text = postion.ToString(CultureInfo.InvariantCulture)
                Me.statusLine.Text = Me.textEditor.GetLineFromCharIndex(pos).ToString(CultureInfo.InvariantCulture)
            ElseIf Me.textInput.Focused Then
                pos = Me.textInput.SelectionStart
                Me.statusPos.Text = pos.ToString(CultureInfo.InvariantCulture)
                Me.statusCol.Text = (pos - Me.textInput.GetFirstCharIndexOfCurrentLine).ToString(CultureInfo.InvariantCulture)
                Me.statusLine.Text = Me.textInput.GetLineFromCharIndex(pos).ToString(CultureInfo.InvariantCulture)
            Else
                Me.statusPos.Text = "-"
                Me.statusCol.Text = "-"
                Me.statusLine.Text = "-"
            End If
        End Sub

        Private Sub tabOutput_Selected(ByVal sender As Object, ByVal e As TabControlEventArgs) Handles tabOutput.Selected
            Me.headerOutput.Text = e.TabPage.Text
        End Sub

        Private Sub TextChangedTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
            Me.TextChangedTimer.Stop()
            Me.textEditor.Invalidate()
            Me.checker.Check(Me.textEditor.Text)
        End Sub

        Private Sub textEditor_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles textEditor.Enter
            Me.SetStatusbar()
        End Sub

        Private Sub textEditor_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles textEditor.Leave
            Me.SetStatusbar()
        End Sub

        Private Sub textEditor_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles textEditor.SelectionChanged
            Me.SetStatusbar()
        End Sub

        Private Sub textEditor_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles textEditor.TextChanged
            Me.marker.Clear()
            Me.TextChangedTimer.Stop()
            Me.TextChangedTimer.Interval = &HBB8
            Me.TextChangedTimer.Start()
            If Not Me.IsDirty Then
                Me.IsDirty = True
                Me.SetFormCaption()
            End If
        End Sub

        Private Sub TextHighlighter_SwitchContext(ByVal sender As Object, ByVal e As ContextSwitchEventArgs)
            If (((((e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.DOTNET_COMMENTBLOCK) OrElse (e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.DOTNET_COMMENTLINE)) OrElse ((e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.DOTNET_STRING) OrElse (e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.GRAMMARSTRING))) OrElse ((e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.DIRECTIVESTRING) OrElse (e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.GRAMMARCOMMENTBLOCK))) OrElse (e.NewContext.TokenField.Type = TinyPG.Highlighter.TokenType.GRAMMARCOMMENTLINE)) Then
                Me.codecomplete.Enabled = False
                Me.directivecomplete.Enabled = False
            ElseIf (e.NewContext.ParentField.TokenField.Type = TokenType.GrammarBlock) Then
                Me.directivecomplete.Enabled = False
                Me.codecomplete.Enabled = True
            ElseIf (e.NewContext.ParentField.TokenField.Type = TokenType.DirectiveBlock) Then
                Me.codecomplete.Enabled = False
                Me.directivecomplete.Enabled = True
            Else
                Me.codecomplete.Enabled = False
                Me.directivecomplete.Enabled = False
            End If
        End Sub

        Private Sub textInput_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles textInput.Enter
            Me.SetStatusbar()
        End Sub

        Private Sub textInput_Leave(ByVal sender As Object, ByVal e As EventArgs) Handles textInput.Leave
            Me.SetStatusbar()
        End Sub

        Private Sub textInput_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles textInput.SelectionChanged
            Me.SetStatusbar()
        End Sub

        Private Sub textOutput_LinkClicked(ByVal sender As Object, ByVal e As LinkClickedEventArgs) Handles textOutput.LinkClicked
            Try
                If (e.LinkText = "www.codeproject.com") Then
                    Process.Start("http://www.codeproject.com/script/Articles/MemberArticles.aspx?amid=2192187")
                End If
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        Private Sub theTinyPGGrammarHighlighterV12ToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles theTinyPGGrammarHighlighterV12ToolStripMenuItem.Click
            MainForm.NotepadViewFile((AppDomain.CurrentDomain.BaseDirectory & "Examples\GrammarHighlighter.tpg"))
        End Sub

        Private Sub theTinyPGGrammarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles theTinyPGGrammarToolStripMenuItem.Click
            MainForm.NotepadViewFile((AppDomain.CurrentDomain.BaseDirectory & "Examples\BNFGrammar v1.1.tpg"))
        End Sub

        Private Sub theTinyPGGrammarV10ToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles theTinyPGGrammarV10ToolStripMenuItem.Click
            MainForm.NotepadViewFile((AppDomain.CurrentDomain.BaseDirectory & "Examples\BNFGrammar v1.0.tpg"))
        End Sub

        Private Sub tvParsetree_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles tvParsetree.AfterSelect
            If (Not e.Node Is Nothing) Then
                Dim ipn As IParseNode = TryCast(e.Node.Tag, IParseNode)
                If (Not ipn Is Nothing) Then
                    Me.textInput.Select(ipn.IToken.StartPos, (ipn.IToken.EndPos - ipn.IToken.StartPos))
                    Me.textInput.ScrollToCaret()
                End If
            End If
        End Sub

        Private Sub ViewFile(ByVal filetype As String)
            Try
                If Not ((Not Me.IsDirty AndAlso (Not Me.compiler Is Nothing)) AndAlso Me.compiler.IsCompiled) Then
                    Me.CompileGrammar()
                End If
                If (Not Me.grammar Is Nothing) Then
                    Dim generator As ICodeGenerator = CodeGeneratorFactory.CreateGenerator(filetype, Me.grammar.Directives.Item("TinyPG").Item("Language"))
                    Process.Start((Me.grammar.GetOutputPath & generator.FileName))
                End If
            Catch exception1 As Exception
            End Try
        End Sub

        Private Sub viewParserToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles viewParserToolStripMenuItem.Click
            Me.ViewFile("Parser")
        End Sub

        Private Sub viewParseTreeCodeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles viewParseTreeCodeToolStripMenuItem.Click
            Me.ViewFile("ParseTree")
        End Sub

        Private Sub viewScannerToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles viewScannerToolStripMenuItem.Click
            Me.ViewFile("Scanner")
        End Sub


        ' Fields
        Private WithEvents aboutTinyParserGeneratorToolStripMenuItem As ToolStripMenuItem
        Private WithEvents checker As SyntaxChecker
        Private WithEvents codeblocksToolStripMenuItem As ToolStripMenuItem
        Private WithEvents codecomplete As AutoComplete
        Private WithEvents compiler As TinyPG.Compiler.Compiler
        Private WithEvents components As IContainer = Nothing
        Private WithEvents directivecomplete As AutoComplete
        Private WithEvents DockExtender As DockExtender
        Private WithEvents examplesToolStripMenuItem As ToolStripMenuItem
        Private WithEvents exitToolStripMenuItem As ToolStripMenuItem
        Private WithEvents expressionEvaluatorToolStripMenuItem As ToolStripMenuItem
        Private WithEvents expressionEvaluatorToolStripMenuItem1 As ToolStripMenuItem
        Private WithEvents fileToolStripMenuItem As ToolStripMenuItem
        Private WithEvents folderBrowserDialog As FolderBrowserDialog
        Private WithEvents grammar As Grammar
        Private WithEvents GrammarFile As String
        Private WithEvents headerEvaluator As HeaderLabel
        Private WithEvents headerOutput As HeaderLabel
        Private WithEvents helpToolStripMenuItem As ToolStripMenuItem
        Private WithEvents highlighterScanner As TinyPG.Highlighter.Scanner
        Private WithEvents inputFloaty As IFloaty
        Private IsDirty As Boolean
        Private WithEvents marker As TextMarker
        Private WithEvents menuStrip As MenuStrip
        Private WithEvents menuToolsGenerate As ToolStripMenuItem
        Private WithEvents newToolStripMenuItem As ToolStripMenuItem
        Private WithEvents openFileDialog As OpenFileDialog
        Private WithEvents openToolStripMenuItem As ToolStripMenuItem
        Private WithEvents outputFloaty As IFloaty
        Private WithEvents outputToolStripMenuItem As ToolStripMenuItem
        Private WithEvents panelInput As Panel
        Private WithEvents panelOutput As Panel
        Private WithEvents parseToolStripMenuItem As ToolStripMenuItem
        Private WithEvents parsetreeToolStripMenuItem As ToolStripMenuItem
        Private WithEvents regExControl As RegExControl
        Private WithEvents regexToolToolStripMenuItem As ToolStripMenuItem
        Private WithEvents saveAsToolStripMenuItem As ToolStripMenuItem
        Private WithEvents saveFileDialog As SaveFileDialog
        Private WithEvents saveToolStripMenuItem As ToolStripMenuItem
        Private WithEvents splitterBottom As Splitter
        Private WithEvents splitterRight As Splitter
        Private WithEvents statusCol As ToolStripStatusLabel
        Private WithEvents statusLabel As ToolStripStatusLabel
        Private WithEvents statusLine As ToolStripStatusLabel
        Private WithEvents statusPos As ToolStripStatusLabel
        Private WithEvents statusStrip As StatusStrip
        Private syntaxUpdateChecker As EventHandler
        Private WithEvents tabOutput As TabControlEx
        Private WithEvents tabPage1 As TabPage
        Private WithEvents tabPage2 As TabPage
        Private WithEvents tabPage3 As TabPage
        Private WithEvents TextChangedTimer As System.Windows.Forms.Timer
        Private WithEvents textEditor As RichTextBox
        Private WithEvents textHighlighter As TextHighlighter
        Private WithEvents textInput As RichTextBox
        Private WithEvents textOutput As RichTextBox
        Private WithEvents theTinyPGGrammarHighlighterV12ToolStripMenuItem As ToolStripMenuItem
        Private WithEvents theTinyPGGrammarToolStripMenuItem As ToolStripMenuItem
        Private WithEvents theTinyPGGrammarV10ToolStripMenuItem As ToolStripMenuItem
        Private WithEvents toolsToolStripMenuItem As ToolStripMenuItem
        Private WithEvents toolStripMenuItem1 As ToolStripSeparator
        Private WithEvents toolStripSeparator1 As ToolStripSeparator
        Private WithEvents toolStripSeparator2 As ToolStripSeparator
        Private WithEvents toolStripStatusLabel1 As ToolStripStatusLabel
        Private WithEvents toolStripStatusLabel2 As ToolStripStatusLabel
        Private WithEvents toolStripStatusLabel3 As ToolStripStatusLabel
        Private WithEvents toolStripStatusLabel4 As ToolStripStatusLabel
        Private WithEvents tvParsetree As TreeView
        Private WithEvents viewParserToolStripMenuItem As ToolStripMenuItem
        Private WithEvents viewParseTreeCodeToolStripMenuItem As ToolStripMenuItem
        Private WithEvents viewScannerToolStripMenuItem As ToolStripMenuItem
        Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
        Friend WithEvents FindToolStripMenuItem As ToolStripMenuItem
        Friend WithEvents Optionals As TabPage
        Friend WithEvents ListBoxOptionals As ListBox
        Private WithEvents viewToolStripMenuItem As ToolStripMenuItem

        Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
            If mfrmFindDialog Is Nothing Then mfrmFindDialog = New DlgFind()
            mfrmFindDialog.mRichTextBox = Me.textEditor
            Me.textInput.HideSelection = False
            mfrmFindDialog.ShowDialog()
            Me.textEditor.Focus()
        End Sub
    End Class
End Namespace

