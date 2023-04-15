Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

Namespace TinyPG.Controls
    Public Class RegExControl
        Inherits UserControl
        ' Methods
        Public Sub New()
            Me.InitializeComponent
        End Sub

        Private Sub checkIgnoreCase_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.ValidateExpression
        End Sub

        Private Sub checkMultiline_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.ValidateExpression
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If (disposing AndAlso (Not Me.components Is Nothing)) Then
                Me.components.Dispose
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            Dim resources As New ComponentResourceManager(GetType(RegExControl))
            Me.Splitter = New Splitter
            Me.panel2 = New Panel
            Me.textMatches = New RichTextBox
            Me.label2 = New Label
            Me.checkIgnoreCase = New CheckBox
            Me.checkMultiline = New CheckBox
            Me.textExpression = New TextBox
            Me.textBox = New RichTextBox
            Me.panel1 = New Panel
            Me.statusText = New Label
            Me.label1 = New Label
            Me.panel2.SuspendLayout
            Me.panel1.SuspendLayout
            MyBase.SuspendLayout
            Me.Splitter.BackColor = SystemColors.InactiveCaption
            Me.Splitter.Dock = DockStyle.Bottom
            Me.Splitter.Location = New Point(0, &H1A5)
            Me.Splitter.Name = "Splitter"
            Me.Splitter.Size = New Size(&H14B, 5)
            Me.Splitter.TabIndex = 11
            Me.Splitter.TabStop = False
            Me.panel2.Controls.Add(Me.textMatches)
            Me.panel2.Controls.Add(Me.label2)
            Me.panel2.Dock = DockStyle.Bottom
            Me.panel2.Location = New Point(0, &H1AA)
            Me.panel2.Name = "panel2"
            Me.panel2.Size = New Size(&H14B, &H39)
            Me.panel2.TabIndex = 12
            Me.textMatches.BackColor = SystemColors.Window
            Me.textMatches.BorderStyle = BorderStyle.None
            Me.textMatches.Dock = DockStyle.Fill
            Me.textMatches.Font = New Font("Courier New", 9.75!)
            Me.textMatches.HideSelection = False
            Me.textMatches.Location = New Point(0, 20)
            Me.textMatches.Name = "textMatches"
            Me.textMatches.ReadOnly = True
            Me.textMatches.Size = New Size(&H14B, &H25)
            Me.textMatches.TabIndex = 6
            Me.textMatches.Text = ""
            Me.label2.BackColor = SystemColors.InactiveBorder
            Me.label2.Dock = DockStyle.Top
            Me.label2.Font = New Font("Segoe UI", 9!)
            Me.label2.Location = New Point(0, 0)
            Me.label2.Name = "label2"
            Me.label2.Size = New Size(&H14B, 20)
            Me.label2.TabIndex = 7
            Me.label2.Text = "Match results (displays group names only)"
            Me.label2.TextAlign = ContentAlignment.MiddleLeft
            Me.checkIgnoreCase.AutoSize = True
            Me.checkIgnoreCase.Checked = True
            Me.checkIgnoreCase.CheckState = CheckState.Checked
            Me.checkIgnoreCase.Font = New Font("Segoe UI", 9!)
            Me.checkIgnoreCase.Location = New Point(8, &H2E)
            Me.checkIgnoreCase.Name = "checkIgnoreCase"
            Me.checkIgnoreCase.Size = New Size(&H56, &H13)
            Me.checkIgnoreCase.TabIndex = 6
            Me.checkIgnoreCase.Text = "Ignore case"
            Me.checkIgnoreCase.UseVisualStyleBackColor = True
            AddHandler Me.checkIgnoreCase.CheckedChanged, New EventHandler(AddressOf Me.checkIgnoreCase_CheckedChanged)
            Me.checkMultiline.AutoSize = True
            Me.checkMultiline.Checked = True
            Me.checkMultiline.CheckState = CheckState.Checked
            Me.checkMultiline.Font = New Font("Segoe UI", 9!)
            Me.checkMultiline.Location = New Point(8, 30)
            Me.checkMultiline.Name = "checkMultiline"
            Me.checkMultiline.Size = New Size(&H49, &H13)
            Me.checkMultiline.TabIndex = 5
            Me.checkMultiline.Text = "Multiline"
            Me.checkMultiline.UseVisualStyleBackColor = True
            AddHandler Me.checkMultiline.CheckedChanged, New EventHandler(AddressOf Me.checkMultiline_CheckedChanged)
            Me.textExpression.Anchor = (AnchorStyles.Right Or (AnchorStyles.Left Or AnchorStyles.Top))
            Me.textExpression.Font = New Font("Microsoft Sans Serif", 12!, FontStyle.Regular, GraphicsUnit.Point, 0)
            Me.textExpression.HideSelection = False
            Me.textExpression.Location = New Point(&H60, 8)
            Me.textExpression.Name = "textExpression"
            Me.textExpression.Size = New Size(220, &H1A)
            Me.textExpression.TabIndex = 1
            AddHandler Me.textExpression.TextChanged, New EventHandler(AddressOf Me.textExpression_TextChanged)
            Me.textBox.BorderStyle = BorderStyle.None
            Me.textBox.Dock = DockStyle.Fill
            Me.textBox.Font = New Font("Courier New", 9.75!, FontStyle.Regular, GraphicsUnit.Point, 0)
            Me.textBox.HideSelection = False
            Me.textBox.Location = New Point(0, &H44)
            Me.textBox.Name = "textBox"
            Me.textBox.Size = New Size(&H14B, &H161)
            Me.textBox.TabIndex = 9
            Me.textBox.Text = resources.GetString("textBox.Text")
            AddHandler Me.textBox.Leave, New EventHandler(AddressOf Me.textBox_Leave)
            Me.panel1.BackColor = SystemColors.InactiveBorder
            Me.panel1.Controls.Add(Me.statusText)
            Me.panel1.Controls.Add(Me.checkIgnoreCase)
            Me.panel1.Controls.Add(Me.checkMultiline)
            Me.panel1.Controls.Add(Me.textExpression)
            Me.panel1.Controls.Add(Me.label1)
            Me.panel1.Dock = DockStyle.Top
            Me.panel1.Location = New Point(0, 0)
            Me.panel1.Name = "panel1"
            Me.panel1.Size = New Size(&H14B, &H44)
            Me.panel1.TabIndex = 8
            Me.statusText.Anchor = (AnchorStyles.Right Or (AnchorStyles.Left Or AnchorStyles.Top))
            Me.statusText.Font = New Font("Segoe UI", 9!)
            Me.statusText.Location = New Point(&H5D, &H23)
            Me.statusText.Name = "statusText"
            Me.statusText.Size = New Size(&HDF, 30)
            Me.statusText.TabIndex = 7
            Me.label1.AutoSize = True
            Me.label1.Font = New Font("Segoe UI", 12!)
            Me.label1.Location = New Point(8, 6)
            Me.label1.Name = "label1"
            Me.label1.Size = New Size(&H57, &H15)
            Me.label1.TabIndex = 4
            Me.label1.Text = "Expression:"
            MyBase.AutoScaleDimensions = New SizeF(6!, 13!)
            MyBase.AutoScaleMode = AutoScaleMode.Font
            MyBase.Controls.Add(Me.textBox)
            MyBase.Controls.Add(Me.Splitter)
            MyBase.Controls.Add(Me.panel2)
            MyBase.Controls.Add(Me.panel1)
            MyBase.Name = "RegExControl"
            MyBase.Size = New Size(&H14B, &H1E3)
            Me.panel2.ResumeLayout(False)
            Me.panel1.ResumeLayout(False)
            Me.panel1.PerformLayout
            MyBase.ResumeLayout(False)
        End Sub

        Private Sub textBox_Leave(ByVal sender As Object, ByVal e As EventArgs)
            Me.ValidateExpression
        End Sub

        Private Sub textBox_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        End Sub

        Private Sub textExpression_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me.ValidateExpression
        End Sub

        Private Sub ValidateExpression()
            Try 
                Me.textBox.SuspendLayout
                Me.textBox.SelectAll
                Me.textBox.SelectionBackColor = Color.White
                Me.textBox.DeselectAll
                Me.textMatches.Text = ""
                Dim options As RegexOptions = If(Me.checkMultiline.Checked, RegexOptions.Multiline, RegexOptions.Singleline)
                If Me.checkIgnoreCase.Checked Then
                    options = (options Or RegexOptions.IgnoreCase)
                End If
                Dim expr As New Regex(Me.textExpression.Text, options)
                Dim ms As MatchCollection = expr.Matches(Me.textBox.Text)
                Me.statusText.Text = (ms.Count & " match(es) found")
                Dim sb As New StringBuilder
                If (ms.Count > 0) Then
                    Dim m As Match
                    For Each m In ms
                        Me.textBox.Select(m.Index, m.Length)
                        Me.textBox.SelectionBackColor = Color.LightPink
                        Dim names As String() = expr.GetGroupNames
                        Dim group As String
                        For Each group In names
                            Dim val As Integer
                            If Not Integer.TryParse(group, val) Then
                                sb.Append(("<" & group & ">="))
                                sb.Append(m.Groups.Item(group).Value)
                                sb.Append(ChrW(13) & ChrW(10))
                            End If
                        Next
                    Next
                End If
                Me.textBox.DeselectAll
                Me.textBox.ResumeLayout
                Me.textMatches.Text = sb.ToString
            Catch ex As Exception
                Me.statusText.Text = ex.Message
            End Try
        End Sub


        ' Fields
        Private checkIgnoreCase As CheckBox
        Private checkMultiline As CheckBox
        Private components As IContainer = Nothing
        Private label1 As Label
        Private label2 As Label
        Private panel1 As Panel
        Private panel2 As Panel
        Private Splitter As Splitter
        Private statusText As Label
        Private textBox As RichTextBox
        Private textExpression As TextBox
        Private textMatches As RichTextBox
    End Class
End Namespace

