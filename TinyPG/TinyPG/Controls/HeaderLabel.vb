Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Namespace TinyPG.Controls
    Friend Class HeaderLabel
        Inherits Label
        ' Events
        Public Event CloseClick As EventHandler

        ' Methods
        Public Sub Activate(ByVal control As Control)
            Me.FocusControl = control
            Me.ActivateRecursive(MyBase.Parent)
        End Sub

        Public Sub ActivatedBy(ByVal control As Control)
            AddHandler control.GotFocus, New EventHandler(AddressOf Me.control_GotFocus)
            AddHandler control.MouseDown, New MouseEventHandler(AddressOf Me.control_MouseDown)
            AddHandler control.LostFocus, New EventHandler(AddressOf Me.control_LostFocus)
        End Sub

        Private Sub ActivateRecursive(ByVal control As Control)
            Me.ActivatedBy(control)
            Dim c As Control
            For Each c In control.Controls
                Me.ActivateRecursive(c)
            Next
        End Sub

        Private Sub control_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
            Me.HasFocus = True
            Me.RefreshHeader()
        End Sub

        Private Sub control_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
            Me.HasFocus = False
            Me.RefreshHeader()
        End Sub

        Private Sub control_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            Me.HasFocus = True
            Me.RefreshHeader()
        End Sub

        Public Sub DeactivatedBy(ByVal control As Control)
            AddHandler control.GotFocus, New EventHandler(AddressOf Me.control_LostFocus)
        End Sub

        Private Function IsHighlighted() As Boolean
            Dim box As New Rectangle((MyBase.Width - 20), (Convert.ToInt32((MyBase.Height - 15) / 2)), &H10, 14)
            Dim p As Point = MyBase.PointToClient(Cursor.Position)
            Dim rect As New Rectangle(p.X, p.Y, 1, 1)
            Return rect.IntersectsWith(box)
        End Function

        Protected Overrides Sub OnCreateControl()
            MyBase.OnCreateControl()
            Me.RefreshHeader()
        End Sub

        Protected Overrides Sub OnGotFocus(ByVal e As EventArgs)
            MyBase.OnGotFocus(e)
            If (Not Me.FocusControl Is Nothing) Then
                Me.FocusControl.Focus()
            Else
                MyBase.Parent.Focus()
            End If
        End Sub

        Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
            If ((e.Button = MouseButtons.Left) AndAlso Me.IsHighlighted) Then
                Me.CloseButtonPressed = True
            Else
                Me.CloseButtonPressed = False
            End If
            MyBase.Focus()
            If Not Me.IsHighlighted Then
                MyBase.OnMouseDown(e)
            End If
        End Sub

        Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
            Me.CloseButtonPressed = False
            MyBase.Invalidate()
            MyBase.OnMouseLeave(e)
        End Sub

        Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
            If Not Me.IsHighlighted Then
                MyBase.OnMouseMove(e)
            End If
            MyBase.Invalidate()
        End Sub

        Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
            If ((Me.CloseButtonPressed AndAlso Me.IsHighlighted) AndAlso (e.Button = MouseButtons.Left)) Then
                RaiseEvent CloseClick(Me, New EventArgs)
            Else
                MyBase.OnMouseUp(e)
            End If
            Me.CloseButtonPressed = False
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            MyBase.OnPaint(e)
            Me.PaintCloseButton(e.Graphics)
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal pevent As PaintEventArgs)
            Dim brush As Brush
            Dim r As New Rectangle(0, 0, MyBase.Width, MyBase.Height)
            If Me.HasFocus Then
                brush = New LinearGradientBrush(r, SystemColors.ActiveCaption, SystemColors.GradientActiveCaption, LinearGradientMode.Vertical)
            Else
                brush = New LinearGradientBrush(r, SystemColors.InactiveCaption, SystemColors.GradientInactiveCaption, LinearGradientMode.Vertical)
            End If
            pevent.Graphics.FillRectangle(brush, r)
            brush.Dispose()
            r.Height -= 1
            r.Width -= 1
            pevent.Graphics.DrawRectangle(SystemPens.ControlDark, r)
        End Sub

        Protected Sub PaintCloseButton(ByVal graphics As Graphics)
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            graphics.InterpolationMode = InterpolationMode.Bicubic
            Dim box As New Rectangle((MyBase.Width - 20), (Convert.ToInt32((MyBase.Height - 15) / 2)), &H10, 14)
            Dim pen As New Pen(If(Me.HasFocus, SystemColors.WindowText, SystemColors.GrayText), 2.0!)
            Dim p1 As New Point((MyBase.Width - &H10), (Convert.ToInt32((MyBase.Height - 8) / 2)))
            Dim p2 As New Point((p1.X + 7), (p1.Y + 7))
            Dim p3 As New Point((p1.X + 7), p1.Y)
            Dim p4 As New Point(p1.X, (p1.Y + 7))
            If Me.IsHighlighted Then
                If Me.CloseButtonPressed Then
                    graphics.FillRectangle(SystemBrushes.GradientInactiveCaption, box)
                    graphics.DrawRectangle(SystemPens.ActiveCaption, box)
                Else
                    graphics.FillRectangle(SystemBrushes.GradientActiveCaption, box)
                    graphics.DrawRectangle(SystemPens.Highlight, box)
                End If
            End If
            graphics.DrawLine(pen, p1, p2)
            graphics.DrawLine(pen, p3, p4)
            pen.Dispose()
        End Sub

        Private Sub RefreshHeader()
            If Me.HasFocus Then
                Me.ForeColor = SystemColors.ControlText
            Else
                Me.ForeColor = SystemColors.GrayText
            End If
            MyBase.Invalidate()
        End Sub


        ' Fields
        Private CloseButtonPressed As Boolean
        Private FocusControl As Control
        Private HasFocus As Boolean
    End Class
End Namespace

