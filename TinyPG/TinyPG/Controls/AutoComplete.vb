Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Namespace TinyPG.Controls
    Public Class AutoComplete
        Inherits Form
        ' Methods
        Public Sub New(ByVal editor As RichTextBox)
            Me.textEditor = editor
            AddHandler Me.textEditor.KeyDown, New KeyEventHandler(AddressOf Me.editor_KeyDown)
            AddHandler Me.textEditor.KeyUp, New KeyEventHandler(AddressOf Me.textEditor_KeyUp)
            AddHandler Me.textEditor.LostFocus, New EventHandler(AddressOf Me.textEditor_LostFocus)
            Me.InitializeComponent
        End Sub

        Private Sub AutoComplete_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
            If ((((e.KeyValue = &H20) OrElse (e.KeyValue = &H1B)) OrElse (e.KeyValue = 13)) OrElse (e.KeyValue = 9)) Then
                MyBase.Visible = False
            End If
            If ((e.KeyValue = 9) OrElse (e.KeyValue = 13)) Then
                Me.SelectCurrentWord
            End If
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If (disposing AndAlso (Not Me.components Is Nothing)) Then
                Me.components.Dispose
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub editor_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If MyBase.Enabled Then
                If (e.KeyValue = &H20) Then
                    If e.Control Then
                        e.Handled = True
                        e.SuppressKeyPress = True
                    End If
                    If (Me.suppress > 0) Then
                        Me.suppress -= 1
                    End If
                End If
                If (e.Control AndAlso (e.KeyValue <> &H20)) Then
                    Me.suppress = 2
                End If
                If ((e.KeyValue = &H1B) AndAlso MyBase.Visible) Then
                    Me.suppress = 2
                End If
                If (MyBase.Visible AndAlso ((((e.KeyValue = &H21) OrElse (e.KeyValue = &H22)) OrElse (e.KeyValue = &H26)) OrElse (e.KeyValue = 40))) Then
                    Me.SendKey(Convert.ToChar(e.KeyValue))
                    e.Handled = True
                End If
            End If
        End Sub

        Private Sub InitializeComponent()
            Me.WordList = New ListBox
            MyBase.SuspendLayout
            Me.WordList.BorderStyle = BorderStyle.FixedSingle
            Me.WordList.Dock = DockStyle.Fill
            Me.WordList.Font = New Font("Segoe UI", 9!)
            Me.WordList.FormattingEnabled = True
            Me.WordList.ItemHeight = 15
            Me.WordList.Location = New Point(0, 0)
            Me.WordList.Name = "WordList"
            Me.WordList.Size = New Size(&H12F, &H89)
            Me.WordList.Sorted = True
            Me.WordList.TabIndex = 0
            Me.WordList.UseTabStops = False
            AddHandler Me.WordList.DoubleClick, New EventHandler(AddressOf Me.WordList_DoubleClick)
            MyBase.AutoScaleDimensions = New SizeF(6!, 13!)
            MyBase.AutoScaleMode = AutoScaleMode.Font
            MyBase.ClientSize = New Size(&H12F, &H8D)
            MyBase.ControlBox = False
            MyBase.Controls.Add(Me.WordList)
            MyBase.FormBorderStyle = FormBorderStyle.SizableToolWindow
            MyBase.KeyPreview = True
            MyBase.Name = "AutoComplete"
            MyBase.ShowIcon = False
            MyBase.ShowInTaskbar = False
            MyBase.StartPosition = FormStartPosition.Manual
            MyBase.TopMost = True
            AddHandler MyBase.KeyUp, New KeyEventHandler(AddressOf Me.AutoComplete_KeyUp)
            MyBase.ResumeLayout(False)
        End Sub

        Protected Overrides Sub OnEnabledChanged(ByVal e As EventArgs)
            MyBase.OnEnabledChanged(e)
            If Not MyBase.Enabled Then
                MyBase.Visible = False
            End If
        End Sub

        Private Sub SelectCurrentWord()
            MyBase.Visible = False
            If (Not Me.WordList.SelectedItem Is Nothing) Then
                Dim temp As Integer = Me.textEditor.SelectionStart
                Me.textEditor.Select(Me.autocompletestart, (temp - Me.autocompletestart))
                Me.textEditor.SelectedText = Me.WordList.SelectedItem.ToString
            End If
        End Sub

        Private Sub SendKey(ByVal key As Char)
            AutoComplete.SendMessage(Me.WordList.Handle, &H100, key, IntPtr.Zero)
        End Sub

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Char, ByVal lParam As IntPtr) As Integer
        End Function

        Private Sub textEditor_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
            If MyBase.Enabled Then
                Try 
                    If ((((e.KeyValue = &H20) AndAlso Not e.Control) OrElse (e.KeyValue = 13)) OrElse (e.KeyValue = &H1B)) Then
                        MyBase.Visible = False
                    ElseIf ((((e.KeyValue > &H40) AndAlso (e.KeyValue < &H5B)) AndAlso Not e.Control) OrElse ((e.KeyValue = &H20) AndAlso e.Control)) Then
                        If Not MyBase.Visible Then
                            Dim line As Integer = Me.textEditor.GetFirstCharIndexOfCurrentLine
                            Dim t As String = Helper.Reverse(Me.textEditor.Text.Substring(line, (Me.textEditor.SelectionStart - line)))
                            Dim i As Integer = t.IndexOfAny((" " & ChrW(13) & ChrW(10) & ChrW(9) & ".;:\/?><-=~`[]{}+!#$%^&*()".ToCharArray).ToCharArray())
                            If (i < 0) Then
                                i = t.Length
                            End If
                            Me.autocompletestart = (Me.textEditor.SelectionStart - i)
                            Me.textEditor.Text.IndexOfAny((" " & ChrW(9) & ChrW(13) & ChrW(10)).ToCharArray)
                            Dim p As Point = Me.textEditor.GetPositionFromCharIndex(Me.autocompletestart)
                            p = Me.textEditor.PointToScreen(p)
                            p.X = (p.X - 8)
                            p.Y = (p.Y + &H16)
                            If ((((Me.textEditor.SelectionStart - Me.autocompletestart) > 0) AndAlso (Me.suppress <= 0)) OrElse ((e.KeyValue = &H20) AndAlso e.Control)) Then
                                MyBase.Location = p
                                MyBase.Visible = MyBase.Enabled
                                Me.textEditor.Focus
                            End If
                        End If
                        Me.WordList.SelectedIndex = Me.WordList.FindString(Me.textEditor.Text.Substring(Me.autocompletestart, (Me.textEditor.SelectionStart - Me.autocompletestart)))
                    ElseIf MyBase.Visible Then
                        If Not ((((e.KeyValue <> 9) OrElse e.Alt) OrElse e.Control) OrElse e.Shift) Then
                            Me.SelectCurrentWord
                            e.Handled = True
                        End If
                        If (Me.textEditor.SelectionStart < Me.autocompletestart) Then
                            MyBase.Visible = False
                        End If
                        If (((((e.KeyValue <> &H21) AndAlso (e.KeyValue <> &H22)) AndAlso (e.KeyValue <> &H26)) AndAlso (e.KeyValue <> 40)) AndAlso MyBase.Visible) Then
                            Me.WordList.SelectedIndex = Me.WordList.FindString(Me.textEditor.Text.Substring(Me.autocompletestart, (Me.textEditor.SelectionStart - Me.autocompletestart)))
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            End If
        End Sub

        Private Sub textEditor_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
            If ((Not Me.textEditor.Focused AndAlso Not Me.Focused) AndAlso Not Me.WordList.Focused) Then
                MyBase.Visible = False
            End If
        End Sub

        Private Sub WordList_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.SelectCurrentWord
        End Sub


        ' Fields
        Private autocompletestart As Integer
        Private components As IContainer = Nothing
        Private suppress As Integer
        Private textEditor As RichTextBox
        Private Const WM_KEYDOWN As Integer = &H100
        Public WordList As ListBox
    End Class
End Namespace

