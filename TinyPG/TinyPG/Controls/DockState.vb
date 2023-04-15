Imports System
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace TinyPG.Controls
    <StructLayout(LayoutKind.Sequential)> _
    Friend Structure DockState
        Public Container As Control
        Public Handle As Control
        Public Splitter As Splitter
        Public OrgDockingParent As Control
        Public OrgDockHost As Control
        Public OrgDockStyle As DockStyle
        Public OrgBounds As Rectangle
    End Structure
End Namespace

