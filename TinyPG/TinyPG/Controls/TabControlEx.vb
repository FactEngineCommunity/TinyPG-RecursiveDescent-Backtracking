Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace TinyPG.Controls
    Friend Class TabControlEx
        Inherits TabControl
        ' Methods
        Protected Overridable Sub DrawTabPage(ByVal graphics As Graphics, ByVal index As Integer, ByVal page As TabPage)
            Dim brush As Brush
            Dim points As Point()
            Dim textwidth As Integer = CInt(graphics.MeasureString(page.Text, Me.Font).Width)
            Dim r As Rectangle = MyBase.GetTabRect(index)
            Dim p As Point = MyBase.PointToClient(Cursor.Position)
            Dim highlight As Boolean = New Rectangle(p.X, p.Y, 1, 1).IntersectsWith(r)
            If (index = MyBase.SelectedIndex) Then
                brush = SystemBrushes.ControlLightLight
                If (MyBase.Alignment = TabAlignment.Top) Then
                    r.X = (r.X - 2)
                    r.Y = (r.Y - 2)
                    r.Width = (r.Width + 2)
                    r.Height = (r.Height + 5)
                Else
                    r.X = (r.X - 2)
                    r.Y = (r.Y - 2)
                    r.Width = (r.Width + 2)
                    r.Height = (r.Height + 4)
                End If
            Else
                If (MyBase.Alignment = TabAlignment.Top) Then
                    r.Y = r.Y
                    r.Height += 1
                Else
                    r.Y = (r.Y - 2)
                    r.Height = (r.Height + 2)
                End If
                If highlight Then
                    brush = New LinearGradientBrush(r, ControlPaint.LightLight(SystemColors.Highlight), SystemColors.ButtonHighlight, LinearGradientMode.Vertical)
                Else
                    brush = New LinearGradientBrush(r, SystemColors.ControlLight, SystemColors.ButtonHighlight, LinearGradientMode.Vertical)
                End If
            End If
            If (MyBase.Alignment = TabAlignment.Top) Then
                graphics.FillRectangle(brush, r)
                points = New Point() { New Point(r.Left, ((r.Top + r.Height) - 1)), New Point(r.Left, r.Top), New Point((r.Left + r.Width), r.Top), New Point((r.Left + r.Width), ((r.Top + r.Height) - 1)) }
                graphics.DrawLines(Pens.Gray, points)
                graphics.DrawString(page.Text, Me.Font, Brushes.Black, CSng((r.Left + ((r.Width - textwidth) / 2))), CSng((r.Top + 2)))
            ElseIf (MyBase.Alignment = TabAlignment.Bottom) Then
                graphics.FillRectangle(brush, CInt((r.Left + 1)), CInt((r.Top + 1)), CInt((r.Width - 1)), CInt((r.Height - 1)))
                points = New Point() { New Point(r.Left, r.Top), New Point(r.Left, ((r.Top + r.Height) - 1)), New Point((r.Left + r.Width), ((r.Top + r.Height) - 1)), New Point((r.Left + r.Width), r.Top) }
                graphics.DrawLines(Pens.Gray, points)
                graphics.DrawString(page.Text, Me.Font, Brushes.Black, CSng((r.Left + ((r.Width - textwidth) / 2))), CSng((r.Top + 2)))
                If (index = MyBase.SelectedIndex) Then
                    graphics.DrawLine(Pens.White, (r.Left + 1), r.Top, ((r.Left + r.Width) - 1), r.Top)
                End If
            End If
        End Sub

        Protected Overrides Sub OnCreateControl()
            MyBase.SetStyle((ControlStyles.DoubleBuffer Or (ControlStyles.AllPaintingInWmPaint Or (ControlStyles.SupportsTransparentBackColor Or ControlStyles.UserPaint))), True)
        End Sub

        Protected Overrides Sub OnMouseEnter(ByVal e As EventArgs)
            MyBase.Invalidate
            MyBase.OnMouseEnter(e)
        End Sub

        Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
            MyBase.Invalidate
            MyBase.OnMouseLeave(e)
        End Sub

        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            If (MyBase.Alignment = TabAlignment.Top) Then
                e.Graphics.FillRectangle(SystemBrushes.ControlLightLight, New Rectangle(0, &H17, (MyBase.Width - 2), (MyBase.Height - 2)))
                e.Graphics.DrawRectangle(SystemPens.ControlDarkDark, New Rectangle(0, &H15, (MyBase.Width - 2), (MyBase.Height - &H17)))
            ElseIf (MyBase.Alignment = TabAlignment.Bottom) Then
                e.Graphics.FillRectangle(SystemBrushes.ControlLightLight, New Rectangle(0, 0, MyBase.Width, (MyBase.Height - 20)))
                e.Graphics.DrawRectangle(SystemPens.ControlDarkDark, New Rectangle(0, 0, (MyBase.Width - 2), (MyBase.Height - &H16)))
            End If
            Dim i As Integer
            For i = 0 To MyBase.TabPages.Count - 1
                If (i <> MyBase.SelectedIndex) Then
                    Dim page As TabPage = MyBase.TabPages.Item(i)
                    Me.DrawTabPage(e.Graphics, i, page)
                End If
            Next i
            If (MyBase.SelectedIndex >= 0) Then
                Me.DrawTabPage(e.Graphics, MyBase.SelectedIndex, MyBase.TabPages.Item(MyBase.SelectedIndex))
            End If
        End Sub

        Protected Overrides Sub OnSystemColorsChanged(ByVal e As EventArgs)
            MyBase.Invalidate
            MyBase.OnSystemColorsChanged(e)
        End Sub

    End Class
End Namespace

