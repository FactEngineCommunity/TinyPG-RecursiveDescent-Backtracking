Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace TinyPG.Controls
    Friend NotInheritable Class Floaty
        Inherits Form
        Implements IFloaty
        ' Events
        Public Custom Event Docking As EventHandler Implements IFloaty.Docking
            AddHandler(ByVal value As EventHandler)

            End AddHandler

            RemoveHandler(ByVal value As EventHandler)

            End RemoveHandler

            RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)

            End RaiseEvent
        End Event

        ' Methods
        Public Sub New(ByVal DockExtender As DockExtender)
            Me._dockExtender = DockExtender
            Me.InitializeComponent()
        End Sub

        Friend Sub Attach(ByVal dockState As DockState)
            Me._dockState = dockState
            Me.Text = Me._dockState.Handle.Text
            AddHandler Me._dockState.Handle.MouseMove, New MouseEventHandler(AddressOf Me.Handle_MouseMove)
            AddHandler Me._dockState.Handle.MouseHover, New EventHandler(AddressOf Me.Handle_MouseHover)
            AddHandler Me._dockState.Handle.MouseLeave, New EventHandler(AddressOf Me.Handle_MouseLeave)
        End Sub

        Private Sub DetachHandle()
            RemoveHandler Me._dockState.Handle.MouseMove, New MouseEventHandler(AddressOf Me.Handle_MouseMove)
            RemoveHandler Me._dockState.Handle.MouseHover, New EventHandler(AddressOf Me.Handle_MouseHover)
            RemoveHandler Me._dockState.Handle.MouseLeave, New EventHandler(AddressOf Me.Handle_MouseLeave)
            Me._dockState.Container = Nothing
            Me._dockState.Handle = Nothing
        End Sub

        Public Shadows Sub Dock() Implements IFloaty.Dock
            If Me._isFloating Then
                Me.DockFloaty()
            End If
        End Sub

        Private Sub DockFloaty()
            Me._dockState.OrgDockHost.TopLevelControl.BringToFront()
            Me.Hide()
            Me._dockState.Container.Visible = False
            Me._dockState.Container.Parent = Me._dockState.OrgDockingParent
            Me._dockState.Container.Dock = Me._dockState.OrgDockStyle
            Me._dockState.Container.Bounds = Me._dockState.OrgBounds
            Me._dockState.Handle.Visible = True
            Me._dockState.Container.Visible = True
            If Me._dockOnInside Then
                Me._dockState.Container.BringToFront()
            End If
            If (((Not Me._dockState.Splitter Is Nothing) AndAlso (Me._dockState.OrgDockStyle <> DockStyle.Fill)) AndAlso (Me._dockState.OrgDockStyle <> DockStyle.None)) Then
                Me._dockState.Splitter.Parent = Me._dockState.OrgDockingParent
                Me._dockState.Splitter.Dock = Me._dockState.OrgDockStyle
                Me._dockState.Splitter.Visible = True
                If Me._dockOnInside Then
                    Me._dockState.Splitter.BringToFront()
                Else
                    Me._dockState.Splitter.SendToBack()
                End If
            End If
            If Not Me._dockOnInside Then
                Me._dockState.Container.SendToBack()
            End If
            Me._isFloating = False

            RaiseEvent Docking(Me, New EventArgs())
        End Sub

        Public Sub Float() Implements IFloaty.Float
            If Not Me._isFloating Then
                Me.Text = Me._dockState.Handle.Text
                Dim pt As Point = Me._dockState.Container.PointToScreen(New Point(0, 0))
                Dim sz As Size = Me._dockState.Container.Size
                If Me._dockState.Container.Equals(Me._dockState.Handle) Then
                    sz.Width = (sz.Width + &H12)
                    sz.Height = (sz.Height + &H1C)
                End If
                If (sz.Width > 600) Then
                    sz.Width = 600
                End If
                If (sz.Height > 600) Then
                    sz.Height = 600
                End If
                Me._dockState.OrgDockingParent = Me._dockState.Container.Parent
                Me._dockState.OrgBounds = Me._dockState.Container.Bounds
                Me._dockState.OrgDockStyle = Me._dockState.Container.Dock
                Me._dockState.Handle.Hide()
                Me._dockState.Container.Parent = Me
                Me._dockState.Container.Dock = DockStyle.Fill
                If (Not Me._dockState.Splitter Is Nothing) Then
                    Me._dockState.Splitter.Visible = False
                    Me._dockState.Splitter.Parent = Me
                End If
                MyBase.Bounds = New Rectangle(pt, sz)
                Me._isFloating = True
                Me.Show()
            End If
        End Sub

        Private Function GetDockingArea(ByVal c As Control) As Rectangle
            Dim r As Rectangle = c.Bounds
            If (Not c.Parent Is Nothing) Then
                r = c.Parent.RectangleToScreen(r)
            End If
            Dim rc As Rectangle = c.ClientRectangle
            Dim borderwidth As Integer = Convert.ToInt32(((r.Width - rc.Width) / 2))
            r.X = (r.X + borderwidth)
            r.Y = (r.Y + ((r.Height - rc.Height) - borderwidth))
            If Not Me._dockOnInside Then
                rc.X = (rc.X + r.X)
                rc.Y = (rc.Y + r.Y)
                Return rc
            End If
            Dim cs As Control
            For Each cs In c.Controls
                If Not cs.Visible Then
                    Continue For
                End If
                Select Case cs.Dock
                    Case DockStyle.Top
                        rc.Y = (rc.Y + cs.Height)
                        rc.Height = (rc.Height - cs.Height)
                        Exit Select
                    Case DockStyle.Bottom
                        rc.Height = (rc.Height - cs.Height)
                        Exit Select
                    Case DockStyle.Left
                        rc.X = (rc.X + cs.Width)
                        rc.Width = (rc.Width - cs.Width)
                        Exit Select
                    Case DockStyle.Right
                        rc.Width = (rc.Width - cs.Width)
                        Exit Select
                End Select
            Next
            rc.X = (rc.X + r.X)
            rc.Y = (rc.Y + r.Y)
            Return rc
        End Function

        Private Sub Handle_MouseHover(ByVal sender As Object, ByVal e As EventArgs)
            Me._startFloating = True
        End Sub

        Private Sub Handle_MouseLeave(ByVal sender As Object, ByVal e As EventArgs)
            Me._startFloating = False
        End Sub

        Private Sub Handle_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            If ((e.Button = MouseButtons.Left) AndAlso Me._startFloating) Then
                Dim ps As Point = Me._dockState.Handle.PointToScreen(New Point(e.X, e.Y))
                Me.MakeFloatable(Me._dockState, e.X, e.Y)
            End If
        End Sub

        Public Shadows Sub Hide() Implements IFloaty.Hide
            If MyBase.Visible Then
                MyBase.Hide()
            End If
            Me._dockState.Container.Hide()
            If (Not Me._dockState.Splitter Is Nothing) Then
                Me._dockState.Splitter.Hide()
            End If
        End Sub

        Private Sub InitializeComponent()
            MyBase.SuspendLayout()
            MyBase.ClientSize = New Size(&HB2, &H7A)
            MyBase.FormBorderStyle = FormBorderStyle.SizableToolWindow
            MyBase.MaximizeBox = False
            MyBase.Name = "Floaty"
            MyBase.ShowIcon = False
            MyBase.ShowInTaskbar = False
            MyBase.StartPosition = FormStartPosition.Manual
            MyBase.ResumeLayout(False)
            Me._dockOnInside = True
            Me._dockOnHostOnly = True
        End Sub

        Private Sub MakeFloatable(ByVal dockState As DockState, ByVal offsetx As Integer, ByVal offsety As Integer)
            Dim ps As Point = Cursor.Position
            Me._dockState = dockState
            Me.Text = Me._dockState.Handle.Text
            Dim sz As Size = Me._dockState.Container.Size
            If Me._dockState.Container.Equals(Me._dockState.Handle) Then
                sz.Width = (sz.Width + &H12)
                sz.Height = (sz.Height + &H1C)
            End If
            If (sz.Width > 600) Then
                sz.Width = 600
            End If
            If (sz.Height > 600) Then
                sz.Height = 600
            End If
            Me._dockState.OrgDockingParent = Me._dockState.Container.Parent
            Me._dockState.OrgBounds = Me._dockState.Container.Bounds
            Me._dockState.OrgDockStyle = Me._dockState.Container.Dock
            Me._dockState.Handle.Hide()
            Me._dockState.Container.Parent = Me
            Me._dockState.Container.Dock = DockStyle.Fill
            If (Not Me._dockState.Splitter Is Nothing) Then
                Me._dockState.Splitter.Visible = False
                Me._dockState.Splitter.Parent = Me
            End If
            Floaty.SendMessage(Me._dockState.Handle.Handle.ToInt32, &H202, 0, 0)
            ps.X = (ps.X - offsetx)
            ps.Y = (ps.Y - offsety)
            MyBase.Bounds = New Rectangle(ps, sz)
            Me._isFloating = True
            Me.Show()
            Floaty.SendMessage(MyBase.Handle.ToInt32, &H112, &HF012, 0)
        End Sub

        Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
            e.Cancel = True
            Me.Hide()
            MyBase.OnClosing(e)
        End Sub

        Protected Overrides Sub OnMove(ByVal e As EventArgs)
            If Not MyBase.IsDisposed Then
                Dim pt As Point = Cursor.Position
                Dim pc As Point = MyBase.PointToClient(pt)
                If (((pc.Y >= -21) AndAlso (pc.Y <= 0)) AndAlso ((pc.X >= -1) AndAlso (pc.X <= MyBase.Width))) Then
                    Dim t As Control = Me._dockExtender.FindDockHost(Me, pt)
                    If (t Is Nothing) Then
                        Me._dockExtender.Overlay.Hide()
                    Else
                        Me.SetOverlay(t, pt)
                    End If
                    MyBase.OnMove(e)
                End If
            End If
        End Sub

        Protected Overrides Sub OnResize(ByVal e As EventArgs)
            MyBase.OnResize(e)
        End Sub

        Protected Overrides Sub OnResizeBegin(ByVal e As EventArgs)
            MyBase.OnResizeBegin(e)
        End Sub

        Protected Overrides Sub OnResizeEnd(ByVal e As EventArgs)
            If (Me._dockExtender.Overlay.Visible AndAlso (Not Me._dockExtender.Overlay.DockHostControl Is Nothing)) Then
                Me._dockState.OrgDockingParent = Me._dockExtender.Overlay.DockHostControl
                Me._dockState.OrgBounds = Me._dockState.Container.RectangleToClient(Me._dockExtender.Overlay.Bounds)
                Me._dockState.OrgDockStyle = Me._dockExtender.Overlay.Dock
                Me._dockExtender.Overlay.Hide()
                Me.DockFloaty()
            End If
            Me._dockExtender.Overlay.DockHostControl = Nothing
            Me._dockExtender.Overlay.Hide()
            MyBase.OnResizeEnd(e)
        End Sub

        <DllImport("User32.dll")> _
        Private Shared Function SendMessage(ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        End Function

        Private Sub SetOverlay(ByVal c As Control, ByVal pc As Point)
            Dim r As Rectangle = Me.GetDockingArea(c)
            Dim rc As Rectangle = r
            Dim rx As Single = (CSng((pc.X - r.Left)) / CSng(r.Width))
            Dim ry As Single = (CSng((pc.Y - r.Top)) / CSng(r.Height))
            Me._dockExtender.Overlay.Dock = DockStyle.None
            If ((((rx > 0.0!) AndAlso (rx < ry)) AndAlso ((rx < 0.25) AndAlso (ry < 0.75))) AndAlso (ry > 0.25)) Then
                r.Width = Convert.ToInt32((r.Width / 2))
                If (r.Width > MyBase.Width) Then
                    r.Width = MyBase.Width
                End If
                Me._dockExtender.Overlay.Dock = DockStyle.Left
            End If
            If ((((rx < 1.0!) AndAlso (rx > ry)) AndAlso ((rx > 0.75) AndAlso (ry < 0.75))) AndAlso (ry > 0.25)) Then
                r.Width = Convert.ToInt32((r.Width / 2))
                If (r.Width > MyBase.Width) Then
                    r.Width = MyBase.Width
                End If
                r.X = ((rc.X + rc.Width) - r.Width)
                Me._dockExtender.Overlay.Dock = DockStyle.Right
            End If
            If ((((ry > 0.0!) AndAlso (ry < rx)) AndAlso ((ry < 0.25) AndAlso (rx < 0.75))) AndAlso (rx > 0.25)) Then
                r.Height = Convert.ToInt32((r.Height / 2))
                If (r.Height > MyBase.Height) Then
                    r.Height = MyBase.Height
                End If
                Me._dockExtender.Overlay.Dock = DockStyle.Top
            End If
            If ((((ry < 1.0!) AndAlso (ry > rx)) AndAlso ((ry > 0.75) AndAlso (rx < 0.75))) AndAlso (rx > 0.25)) Then
                r.Height = Convert.ToInt32((r.Height / 2))
                If (r.Height > MyBase.Height) Then
                    r.Height = MyBase.Height
                End If
                r.Y = ((rc.Y + rc.Height) - r.Height)
                Me._dockExtender.Overlay.Dock = DockStyle.Bottom
            End If
            If (Me._dockExtender.Overlay.Dock <> DockStyle.None) Then
                Me._dockExtender.Overlay.Bounds = r
            Else
                Me._dockExtender.Overlay.Hide()
            End If
            If Not (Me._dockExtender.Overlay.Visible OrElse (Me._dockExtender.Overlay.Dock = DockStyle.None)) Then
                Me._dockExtender.Overlay.DockHostControl = c
                Me._dockExtender.Overlay.Show(Me._dockState.OrgDockHost)
                MyBase.BringToFront()
            End If
        End Sub

        Public Shadows Sub Show() Implements IFloaty.Show
            If Not (MyBase.Visible OrElse Not Me._isFloating) Then
                MyBase.Show(Me._dockState.OrgDockHost)
            End If
            Me._dockState.Container.Show()
            If (Not Me._dockState.Splitter Is Nothing) Then
                Me._dockState.Splitter.Show()
            End If
        End Sub

        Public Shadows Sub Show(ByVal win As IWin32Window)
            Me.Show()
        End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)
            If (m.Msg = &HA3) Then
                Me.DockFloaty()
            End If
            MyBase.WndProc((m))
        End Sub


        ' Properties
        Public Property DockOnHostOnly As Boolean Implements IFloaty.DockOnHostOnly
            Get
                Return Me._dockOnHostOnly
            End Get
            Set(ByVal value As Boolean)
                Me._dockOnHostOnly = value
            End Set
        End Property

        Public Property DockOnInside As Boolean Implements IFloaty.DockOnInside
            Get
                Return Me._dockOnInside
            End Get
            Set(ByVal value As Boolean)
                Me._dockOnInside = value
            End Set
        End Property

        Friend ReadOnly Property DockState As DockState
            Get
                Return Me._dockState
            End Get
        End Property

        Public Shadows Property Text As String Implements IFloaty.Text
            Get
                Return Me.textField
            End Get
            Set(ByVal value As String)
                Me.textField = value
            End Set
        End Property

        ' Fields
        Private _dockExtender As DockExtender
        Private _dockOnHostOnly As Boolean
        Private _dockOnInside As Boolean
        Private _dockState As DockState
        Private _isFloating As Boolean
        Private _startFloating As Boolean
        Private textField As String

        Private Const SC_MOVE As Integer = &HF010
        Private Const WM_LBUTTONUP As Integer = &H202
        Private Const WM_NCLBUTTONDBLCLK As Integer = &HA3
        Private Const WM_SYSCOMMAND As Integer = &H112
    End Class
End Namespace

