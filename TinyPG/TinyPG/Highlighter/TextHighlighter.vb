Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Namespace TinyPG.Highlighter
    Public Class TextHighlighter
        Implements IDisposable
        ' Events
        Public Event SwitchContext As ContextSwitchEventHandler

        ' Methods
        Public Sub New(ByVal textbox As RichTextBox, ByVal scanner As Scanner, ByVal parser As Parser)
            Me.Textbox = textbox
            Me.Scanner = scanner
            Me.Parser = parser
            Me.ClearUndo()
            AddHandler Me.Textbox.TextChanged, New EventHandler(AddressOf Me.Textbox_TextChanged)
            AddHandler textbox.KeyDown, New KeyEventHandler(AddressOf Me.textbox_KeyDown)
            AddHandler Me.Textbox.SelectionChanged, New EventHandler(AddressOf Me.Textbox_SelectionChanged)
            AddHandler Me.Textbox.Disposed, New EventHandler(AddressOf Me.Textbox_Disposed)

            Me.currentContext = Me.Tree
            Me.threadAutoHighlight = New Thread(New ThreadStart(AddressOf Me.AutoHighlightStart))
            Me.threadAutoHighlight.Start()
        End Sub

        Private Sub AddRtfEnd(ByVal sb As StringBuilder)
            sb.Append("}")
        End Sub

        Private Sub AddRtfHeader(ByVal sb As StringBuilder)
            sb.Insert(0, "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Consolas;}}{\colortbl;\red0\green128\blue0;\red0\green128\blue0;\red255\green0\blue0;\red128\green0\blue255;\red128\green0\blue128;\red128\green0\blue128;\red43\green145\blue202;\red0\green0\blue255;\red255\green0\blue0;\red0\green0\blue255;\red43\green145\blue202;\red0\green128\blue0;\red0\green128\blue0;\red163\green21\blue21;\red0\green128\blue0;\red0\green128\blue0;\red163\green21\blue21;\red0\green128\blue0;\red0\green128\blue0;\red163\green21\blue21;\red128\green0\blue128;\red128\green0\blue128;\red0\green0\blue255;\red128\green0\blue128;\red163\green21\blue21;}\viewkind4\uc1\pard\lang1033\f0\fs20")
        End Sub

        Private Sub AutoHighlightStart()
            Dim _currenttext As String = ""
            Do While Not Me.isDisposing
                Dim _textchanged As Boolean
                Dim lockObj As Object = TextHighlighter.treelock
                SyncLock lockObj
                    _textchanged = Me.textChanged
                    If Me.textChanged Then
                        Me.textChanged = False
                        _currenttext = Me.currentText
                    End If
                End SyncLock
                If Not _textchanged Then
                    Thread.Sleep(200)
                Else
                    Dim _tree As ParseTree = Me.Parser.Parse(_currenttext)
                    SyncLock lockObj
                        If Me.textChanged Then
                            Continue Do
                        End If
                        Me.Tree = _tree
                    End SyncLock
                    Me.Textbox.Invoke(New MethodInvoker(AddressOf Me.HighlightTextInternal))
                End If
            Loop
        End Sub

        Public Sub ClearUndo()
            Me.UndoList = New List(Of UndoItem)
            Me.UndoIndex = 0
        End Sub

        Sub Dispose() Implements System.IDisposable.Dispose
            Me.isDisposing = True
            Me.threadAutoHighlight.Join(&H3E8)
            If Me.threadAutoHighlight.IsAlive Then
                Me.threadAutoHighlight.Abort()
            End If
        End Sub

        Private Sub [Do](ByVal [text] As String, ByVal position As Integer)
            If (Me.stateLocked = IntPtr.Zero) Then
                Dim ua As New UndoItem([text], position, New Point(Me.HScrollPos, Me.VScrollPos))
                Me.UndoList.RemoveRange(Me.UndoIndex, (Me.UndoList.Count - Me.UndoIndex))
                Me.UndoList.Add(ua)
                If (Me.UndoList.Count > 100) Then
                    Me.UndoList.RemoveAt(0)
                End If
                If (Me.UndoList.Count > 7) Then
                    Dim canRemove As Boolean = True
                    Dim nextItem As UndoItem = ua
                    Dim i As Integer
                    For i = 0 To 6 - 1
                        Dim prevItem As UndoItem = Me.UndoList.Item(((Me.UndoList.Count - 2) - i))
                        canRemove = (canRemove And ((Math.Abs(CInt((prevItem.Text.Length - nextItem.Text.Length))) <= 1) AndAlso (Math.Abs(CInt((prevItem.Position - nextItem.Position))) <= 1)))
                        nextItem = prevItem
                    Next i
                    If canRemove Then
                        Me.UndoList.RemoveRange((Me.UndoList.Count - 6), 5)
                    End If
                End If
                Me.UndoIndex = Me.UndoList.Count
            End If
        End Sub

        Private Function FindNode(ByVal node As ParseNode, ByVal posstart As Integer) As ParseNode
            If ((Not node Is Nothing) AndAlso ((node.TokenField.StartPos <= posstart) AndAlso ((node.TokenField.StartPos + node.TokenField.Length) >= posstart))) Then
                Dim n As ParseNode
                For Each n In node.Nodes
                    If ((n.TokenField.StartPos <= posstart) AndAlso ((n.TokenField.StartPos + n.TokenField.Length) >= posstart)) Then
                        Return Me.FindNode(n, posstart)
                    End If
                Next
                Return node
            End If
            Return Nothing
        End Function

        Public Function GetCurrentContext() As ParseNode
            Return Me.FindNode(Me.Tree, Me.Textbox.SelectionStart)
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function GetScrollPos(ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
        End Function

        Private Sub HighlighTextCore()
            Dim sb As New StringBuilder
            If (Not Me.Tree Is Nothing) Then
                Dim start As ParseNode = Me.Tree.Nodes.Item(0)
                Me.HightlightNode(start, sb)
                Dim skiptoken As Token
                For Each skiptoken In Me.Scanner.Skipped
                    Me.HighlightToken(skiptoken, sb)
                    sb.Append(skiptoken.Text.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}").Replace(ChrW(10), "\par" & ChrW(10)))
                Next
                sb = Me.Unicode(sb)
                Me.AddRtfHeader(sb)
                Me.AddRtfEnd(sb)
                Me.Textbox.Rtf = sb.ToString
            End If
        End Sub

        Public Sub HighlightText()
            SyncLock TextHighlighter.treelock
                Me.textChanged = True
                Me.currentText = Me.Textbox.Text
            End SyncLock
        End Sub

        Private Sub HighlightTextInternal()
            Me.Lock()
            Dim hscroll As Integer = Me.HScrollPos
            Dim vscroll As Integer = Me.VScrollPos
            Dim selstart As Integer = Me.Textbox.SelectionStart
            Me.HighlighTextCore()
            Me.Textbox.Select(selstart, 0)
            Me.HScrollPos = hscroll
            Me.VScrollPos = vscroll
            Me.Unlock()
        End Sub

        Private Sub HighlightToken(ByVal token As Token, ByVal sb As StringBuilder)
            Select Case token.Type
                Case TokenType.GRAMMARCOMMENTLINE
                    sb.Append("{{\cf1 ")
                    Return
                Case TokenType.GRAMMARCOMMENTBLOCK
                    sb.Append("{{\cf2 ")
                    Return
                Case TokenType.DIRECTIVESTRING
                    sb.Append("{{\cf3 ")
                    Return
                Case TokenType.DIRECTIVEKEYWORD
                    sb.Append("{{\cf4 ")
                    Return
                Case TokenType.DIRECTIVEOPEN
                    sb.Append("{{\cf5 ")
                    Return
                Case TokenType.DIRECTIVECLOSE
                    sb.Append("{{\cf6 ")
                    Return
                Case TokenType.ATTRIBUTEKEYWORD
                    sb.Append("{{\cf7 ")
                    Return
                Case TokenType.CS_KEYWORD
                    sb.Append("{{\cf8 ")
                    Return
                Case TokenType.VB_KEYWORD
                    sb.Append("{{\cf9 ")
                    Return
                Case TokenType.DOTNET_KEYWORD
                    sb.Append("{{\cf10 ")
                    Return
                Case TokenType.DOTNET_TYPES
                    sb.Append("{{\cf11 ")
                    Return
                Case TokenType.CS_COMMENTLINE
                    sb.Append("{{\cf12 ")
                    Return
                Case TokenType.CS_COMMENTBLOCK
                    sb.Append("{{\cf13 ")
                    Return
                Case TokenType.CS_STRING
                    sb.Append("{{\cf14 ")
                    Return
                Case TokenType.VB_COMMENTLINE
                    sb.Append("{{\cf15 ")
                    Return
                Case TokenType.VB_COMMENTBLOCK
                    sb.Append("{{\cf16 ")
                    Return
                Case TokenType.VB_STRING
                    sb.Append("{{\cf17 ")
                    Return
                Case TokenType.DOTNET_COMMENTLINE
                    sb.Append("{{\cf18 ")
                    Return
                Case TokenType.DOTNET_COMMENTBLOCK
                    sb.Append("{{\cf19 ")
                    Return
                Case TokenType.DOTNET_STRING
                    sb.Append("{{\cf20 ")
                    Return
                Case TokenType.CODEBLOCKOPEN
                    sb.Append("{{\cf21 ")
                    Return
                Case TokenType.CODEBLOCKCLOSE
                    sb.Append("{{\cf22 ")
                    Return
                Case TokenType.GRAMMARKEYWORD
                    sb.Append("{{\cf23 ")
                    Return
                Case TokenType.GRAMMARARROW
                    sb.Append("{{\cf24 ")
                    Return
                Case TokenType.GRAMMARSTRING
                    sb.Append("{{\cf25 ")
                    Return
            End Select
            sb.Append("{{\cf0 ")
        End Sub

        Private Sub HightlightNode(ByVal node As ParseNode, ByVal sb As StringBuilder)
            If (node.Nodes.Count = 0) Then
                If (Not node.TokenField.Skipped Is Nothing) Then
                    Dim skiptoken As Token
                    For Each skiptoken In node.TokenField.Skipped
                        Me.HighlightToken(skiptoken, sb)
                        sb.Append(skiptoken.Text.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}").Replace(ChrW(10), "\par" & ChrW(10)))
                    Next
                End If
                Me.HighlightToken(node.TokenField, sb)
                sb.Append(node.TokenField.Text.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}").Replace(ChrW(10), "\par" & ChrW(10)))
                sb.Append("}")
            End If
            Dim n As ParseNode
            For Each n In node.Nodes
                Me.HightlightNode(n, sb)
            Next
        End Sub

        Public Sub Lock()
            TextHighlighter.SendMessage(Me.Textbox.Handle, 11, 0, IntPtr.Zero)
            Me.stateLocked = TextHighlighter.SendMessage(Me.Textbox.Handle, &H43B, 0, IntPtr.Zero)
        End Sub

        <DllImport("user32.dll")> _
        Private Shared Function PostMessageA(ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Boolean
        End Function

        Public Sub Redo()
            If Me.CanRedo Then
                Me.UndoIndex += 1
                If (Me.UndoIndex > Me.UndoList.Count) Then
                    Me.UndoIndex = Me.UndoList.Count
                End If
                Dim ua As UndoItem = Me.UndoList.Item((Me.UndoIndex - 1))
                Me.RestoreState(ua)
            End If
        End Sub

        Private Sub RestoreState(ByVal item As UndoItem)
            Me.Lock()
            Me.Textbox.Rtf = item.Text
            Me.Textbox.Select(item.Position, 0)
            Me.HScrollPos = item.ScrollPosition.X
            Me.VScrollPos = item.ScrollPosition.Y
            Me.Unlock()
        End Sub

        <DllImport("user32", CharSet:=CharSet.Auto)> _
        Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll")> _
        Private Shared Function SetScrollPos(ByVal hWnd As IntPtr, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Boolean) As Integer
        End Function

        Private Sub Textbox_Disposed(ByVal sender As Object, ByVal e As EventArgs)
            Me.Dispose()
        End Sub

        Private Sub textbox_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If ((e.KeyValue = &H59) AndAlso e.Control) Then
                Me.Redo()
            End If
            If ((e.KeyValue = 90) AndAlso e.Control) Then
                Me.Undo()
            End If
        End Sub

        Private Sub Textbox_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            If ((Me.stateLocked = IntPtr.Zero)) Then
                Dim newContext As ParseNode = Me.GetCurrentContext
                If (Me.currentContext Is Nothing) Then
                    Me.currentContext = newContext
                End If
                If ((Not newContext Is Nothing) AndAlso (newContext.TokenField.Type <> Me.currentContext.TokenField.Type)) Then
                    RaiseEvent SwitchContext(Me, New ContextSwitchEventArgs(Me.currentContext, newContext))
                    Me.currentContext = newContext
                End If
            End If
        End Sub

        Private Sub Textbox_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
            If Not (Me.stateLocked <> IntPtr.Zero) Then
                Me.Do(Me.Textbox.Rtf, Me.Textbox.SelectionStart)
                Me.HighlightText()
            End If
        End Sub

        Public Sub Undo()
            If Me.CanUndo Then
                Me.UndoIndex -= 1
                If (Me.UndoIndex < 1) Then
                    Me.UndoIndex = 1
                End If
                Dim ua As UndoItem = Me.UndoList.Item((Me.UndoIndex - 1))
                Me.RestoreState(ua)
            End If
        End Sub

        Public Function [Unicode](ByVal sb As StringBuilder) As StringBuilder
            Dim i As Integer = 0
            Dim uc As New StringBuilder

            Do While (i <= (sb.Length - 1))
                Dim c As Char = sb.Chars(i)
                If (c < ChrW(127)) Then
                    uc.Append(c)
                Else
                    uc.Append(("\u" & Convert.ToInt32(c).ToString & "?"))
                End If
                i += 1
            Loop
            Return uc
        End Function

        Public Sub Unlock()
            TextHighlighter.SendMessage(Me.Textbox.Handle, &H445, 0, Me.stateLocked)
            TextHighlighter.SendMessage(Me.Textbox.Handle, 11, 1, IntPtr.Zero)
            Me.stateLocked = IntPtr.Zero
            Me.Textbox.Invalidate()
        End Sub


        ' Properties
        Public ReadOnly Property CanRedo As Boolean
            Get
                Return (Me.UndoIndex < Me.UndoList.Count)
            End Get
        End Property

        Public ReadOnly Property CanUndo As Boolean
            Get
                Return (Me.UndoIndex > 0)
            End Get
        End Property

        Private Property HScrollPos As Integer
            Get
                Return TextHighlighter.GetScrollPos(CInt(Me.Textbox.Handle), 0)
            End Get
            Set(ByVal value As Integer)
                TextHighlighter.SetScrollPos(Me.Textbox.Handle, 0, value, True)
                TextHighlighter.PostMessageA(Me.Textbox.Handle, &H114, (4 + (&H10000 * value)), 0)
            End Set
        End Property

        Private Property VScrollPos As Integer
            Get
                Return TextHighlighter.GetScrollPos(CInt(Me.Textbox.Handle), 1)
            End Get
            Set(ByVal value As Integer)
                TextHighlighter.SetScrollPos(Me.Textbox.Handle, 1, value, True)
                TextHighlighter.PostMessageA(Me.Textbox.Handle, &H115, (4 + (&H10000 * value)), 0)
            End Set
        End Property


        ' Fields
        Private currentContext As ParseNode
        Private currentText As String
        Private Const EM_GETEVENTMASK As Integer = &H43B
        Private Const EM_SETEVENTMASK As Integer = &H445
        Private isDisposing As Boolean
        Private Parser As Parser
        Private Const SB_HORZ As Integer = 0
        Private Const SB_THUMBPOSITION As Integer = 4
        Private Const SB_VERT As Integer = 1
        Private Scanner As Scanner
        Private stateLocked As IntPtr = IntPtr.Zero
        Public ReadOnly Textbox As RichTextBox
        Private textChanged As Boolean
        Private threadAutoHighlight As Thread
        Public Tree As ParseTree
        Public Shared treelock As Object = New Object
        Private Const UNDO_BUFFER As Integer = 100
        Private UndoIndex As Integer = -1
        Private UndoList As List(Of UndoItem)
        Private Const WM_HSCROLL As Integer = &H114
        Private Const WM_SETREDRAW As Integer = 11
        Private Const WM_USER As Integer = &H400
        Private Const WM_VSCROLL As Integer = &H115

        ' Nested Types
        Private Class UndoItem
            ' Methods
            Public Sub New(ByVal [text] As String, ByVal position As Integer, ByVal scroll As Point)
                Me.Text = [text]
                Me.Position = position
                Me.ScrollPosition = scroll
            End Sub


            ' Fields
            Public Position As Integer
            Public ScrollPosition As Point
            Public [Text] As String
        End Class
    End Class
End Namespace

