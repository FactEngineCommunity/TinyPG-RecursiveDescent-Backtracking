Imports System
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports TinyPG.Compiler
Imports TinyPG.Controls

Namespace TinyPG
    Public NotInheritable Class SyntaxChecker
        Implements IDisposable
        ' Events
        Public Event UpdateSyntax As EventHandler

        Private disposingField As Boolean
        Private markerField As TextMarker
        Private textField As String
        Private textchangedField As Boolean
        Private grammarField As Grammar
        Private syntaxTreeField As ParseTree

        Public Property Grammar As Grammar
            Get
                Return Me.grammarField
            End Get
            Set(ByVal value As Grammar)
                Me.grammarField = value
            End Set
        End Property

        Public Property SyntaxTree As ParseTree
            Get
                Return Me.syntaxTreeField
            End Get
            Set(ByVal value As ParseTree)
                Me.syntaxTreeField = value
            End Set
        End Property

        Public Sub New(ByVal marker As TextMarker)
            Me.markerField = marker
            Me.disposingField = False
        End Sub

        Public Sub Check(ByVal [text] As String)
            Me.textField = [text]
            Me.textchangedField = True
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Me.disposingField = True
        End Sub

        Public Sub Start()
            Dim scanner As New Scanner
            Dim parser As New Parser(scanner)
            Do While Not Me.disposingField
                Thread.Sleep(250)
                If Me.textchangedField Then
                    Me.textchangedField = False
                    scanner.Init(Me.textField)
                    Me.SyntaxTree = parser.Parse(Me.textField, New GrammarTree)
                    If (Me.SyntaxTree.Errors.Count > 0) Then
                        Me.SyntaxTree.Errors.Clear()
                    End If
                    Try
                        If (Me.Grammar Is Nothing) Then
                            Me.Grammar = DirectCast(Me.SyntaxTree.Eval(New Object(0 - 1) {}), Grammar)
                        Else
                            SyncLock Me.Grammar
                                Me.Grammar = DirectCast(Me.SyntaxTree.Eval(New Object(0 - 1) {}), Grammar)
                            End SyncLock
                        End If
                    Catch exception1 As Exception
                    End Try
                    If Not Me.textchangedField Then
                        SyncLock Me.markerField
                            Me.markerField.Clear()
                            Dim err As ParseError
                            For Each err In Me.SyntaxTree.Errors
                                Me.markerField.AddWord(err.Position, err.Length, Color.Red, err.Message)
                            Next
                        End SyncLock
                        RaiseEvent UpdateSyntax(Me, New EventArgs)
                    End If
                End If
            Loop
        End Sub

    End Class
End Namespace

