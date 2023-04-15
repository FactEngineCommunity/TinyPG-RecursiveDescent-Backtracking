Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace TinyPG.Controls
    Friend NotInheritable Class Overlay
        Inherits Form
        ' Methods
        Public Sub New()
            Me.InitializeComponent
        End Sub

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If (disposing AndAlso (Not Me.components Is Nothing)) Then
                Me.components.Dispose
            End If
            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            MyBase.SuspendLayout
            Me.BackColor = SystemColors.ActiveCaption
            MyBase.ControlBox = False
            MyBase.FormBorderStyle = FormBorderStyle.None
            MyBase.MaximizeBox = False
            MyBase.MinimizeBox = False
            MyBase.Name = "Overlay"
            MyBase.Opacity = 0.3
            MyBase.ShowIcon = False
            MyBase.ShowInTaskbar = False
            MyBase.StartPosition = FormStartPosition.Manual
            Me.Text = "Overlay"
            MyBase.ResumeLayout(False)
        End Sub


        ' Fields
        Private components As IContainer = Nothing
        Public Shadows Dock As DockStyle
        Public DockHostControl As Control
    End Class
End Namespace

