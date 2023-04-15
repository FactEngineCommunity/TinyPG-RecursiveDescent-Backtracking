Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace TinyPG.Controls
    <ProvideProperty("Dockable", GetType(Panel))> _
    Public NotInheritable Class DockExtender
        Inherits Component
        Implements IExtenderProvider, ISupportInitialize
        ' Methods
        Public Sub New()
            Me.Overlay = New Overlay
            Me._dockHost = Nothing
            Me._floaties = New Floaties
        End Sub

        Public Sub New(ByVal dockHost As Control)
            Me.Overlay = New Overlay
            Me._dockHost = dockHost
            Me._floaties = New Floaties
        End Sub

        Public Function Attach(ByVal container As Control) As IFloaty
            Return Me.Attach(container, container, Nothing)
        End Function

        Public Function Attach(ByVal container As Control, ByVal handle As Control) As IFloaty
            Return Me.Attach(container, handle, Nothing)
        End Function

        Public Function Attach(ByVal container As Control, ByVal handle As Control, ByVal splitter As Splitter) As IFloaty
            If (container Is Nothing) Then
                Throw New ArgumentException("container cannot be null")
            End If
            If (handle Is Nothing) Then
                Throw New ArgumentException("handle cannot be null")
            End If
            Dim _dockState As New DockState With { _
                .Container = container, _
                .Handle = handle, _
                .OrgDockHost = Me._dockHost, _
                .Splitter = splitter _
            }
            Dim floaty As New Floaty(Me)
            floaty.Attach(_dockState)
            Me._floaties.Add(floaty)
            Return floaty
        End Function

        Public Sub BeginInit() Implements ISupportInitialize.BeginInit
            Console.WriteLine("DockExtender_BeginInit")
        End Sub

        Public Function CanExtend(ByVal extendee As Object) As Boolean Implements IExtenderProvider.CanExtend
            Return TypeOf extendee Is Control
        End Function

        Public Sub EndInit() Implements ISupportInitialize.EndInit
            Console.WriteLine("DockExtender_EndInit")
        End Sub

        Friend Function FindDockHost(ByVal floaty As Floaty, ByVal pt As Point) As Control
            Dim c As Control = Nothing
            If Me.FormIsHit(floaty.DockState.OrgDockHost, pt) Then
                c = floaty.DockState.OrgDockHost
            End If
            If Not floaty.DockOnHostOnly Then
                Dim f As Floaty
                For Each f In Me.Floaties
                    If (f.DockState.Container.Visible AndAlso Me.FormIsHit(f.DockState.Container, pt)) Then
                        Return f.DockState.Container
                    End If
                Next
            End If
            Return c
        End Function

        Friend Function FormIsHit(ByVal c As Control, ByVal pt As Point) As Boolean
            If (c Is Nothing) Then
                Return False
            End If
            Dim pc As Point = c.PointToClient(pt)
            Return c.ClientRectangle.IntersectsWith(New Rectangle(pc, New Size(1, 1)))
        End Function

        Public Sub Hide(ByVal container As Control)
            Dim f As IFloaty = Me._floaties.Find(container)
            If (Not f Is Nothing) Then
                f.Hide
            End If
        End Sub

        Public Sub Show(ByVal container As Control)
            Dim f As IFloaty = Me._floaties.Find(container)
            If (Not f Is Nothing) Then
                f.Show
            End If
        End Sub


        ' Properties
        Public Property Dockable As Boolean
            Get
                Return Me._dockable
            End Get
            Set(ByVal value As Boolean)
                Me._dockable = value
            End Set
        End Property

        Public ReadOnly Property Floaties As Floaties
            Get
                Return Me._floaties
            End Get
        End Property


        ' Fields
        Private _dockable As Boolean
        Private _dockHost As Control
        Private _floaties As Floaties
        Friend Overlay As Overlay
    End Class
End Namespace

