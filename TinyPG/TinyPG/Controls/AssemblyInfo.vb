Imports System
Imports System.Reflection

Namespace TinyPG.Controls
    Public Class AssemblyInfo

        ' Fields
        Private Shared companyNameField As String = String.Empty
        Private Shared copyRightsDetailField As String = String.Empty
        Private Shared productDescriptionField As String = String.Empty
        Private Shared productNameField As String = String.Empty
        Private Shared productTitleField As String = String.Empty
        Private Shared productVersionField As String = String.Empty

        ' Properties
        Public Shared ReadOnly Property CompanyName As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyCompanyAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.companyNameField = DirectCast(customAttributes(0), AssemblyCompanyAttribute).Company
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.companyNameField) Then
                        AssemblyInfo.companyNameField = String.Empty
                    End If
                End If
                Return AssemblyInfo.companyNameField
            End Get
        End Property

        Public Shared ReadOnly Property CopyRightsDetail As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyCopyrightAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.copyRightsDetailField = DirectCast(customAttributes(0), AssemblyCopyrightAttribute).Copyright
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.copyRightsDetailField) Then
                        AssemblyInfo.copyRightsDetailField = String.Empty
                    End If
                End If
                Return AssemblyInfo.copyRightsDetailField
            End Get
        End Property

        Public Shared ReadOnly Property ProductDescription As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyDescriptionAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.productDescriptionField = DirectCast(customAttributes(0), AssemblyDescriptionAttribute).Description
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.productDescriptionField) Then
                        AssemblyInfo.productDescriptionField = String.Empty
                    End If
                End If
                Return AssemblyInfo.productDescriptionField
            End Get
        End Property

        Public Shared ReadOnly Property ProductName As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyProductAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.productNameField = DirectCast(customAttributes(0), AssemblyProductAttribute).Product
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.productNameField) Then
                        AssemblyInfo.productNameField = String.Empty
                    End If
                End If
                Return AssemblyInfo.productNameField
            End Get
        End Property

        Public Shared ReadOnly Property ProductTitle As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyTitleAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.productTitleField = DirectCast(customAttributes(0), AssemblyTitleAttribute).Title
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.productTitleField) Then
                        AssemblyInfo.productTitleField = String.Empty
                    End If
                End If
                Return AssemblyInfo.productTitleField
            End Get
        End Property

        Public Shared ReadOnly Property ProductVersion As String
            Get
                Dim assembly As Assembly = assembly.GetEntryAssembly
                If (Not [assembly] Is Nothing) Then
                    Dim customAttributes As Object() = [assembly].GetCustomAttributes(GetType(AssemblyVersionAttribute), False)
                    If ((Not customAttributes Is Nothing) AndAlso (customAttributes.Length > 0)) Then
                        AssemblyInfo.productVersionField = DirectCast(customAttributes(0), AssemblyVersionAttribute).Version
                    End If
                    If String.IsNullOrEmpty(AssemblyInfo.productVersionField) Then
                        AssemblyInfo.productVersionField = String.Empty
                    End If
                End If
                Return AssemblyInfo.productVersionField
            End Get
        End Property

    End Class
End Namespace

