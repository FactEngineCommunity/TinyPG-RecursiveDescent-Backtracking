Imports System
Imports System.Runtime.CompilerServices

Namespace TinyPG.Controls
    Public Interface IFloaty
        ' Events
        Event Docking As EventHandler

        ' Methods
        Sub Dock()
        Sub Float()
        Sub Hide()
        Sub Show()

        ' Properties
        Property DockOnHostOnly As Boolean

        Property DockOnInside As Boolean
            
        Property [Text] As String
            
    End Interface
End Namespace

