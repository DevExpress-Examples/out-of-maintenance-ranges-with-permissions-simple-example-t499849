Imports System.Collections.Generic
Imports DevExpress.Xpf.Ribbon
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Services
Imports System.Windows.Media

Namespace RangePermissionsSimpleExample
    Partial Public Class MainWindow
        Inherits DXRibbonWindow

        Public Sub New()
            InitializeComponent()
#Region "#RegisterUserList"
            richEdit.ReplaceService(Of IUserListService)(New MyUserListService())
#End Region ' #RegisterUserList

#Region "#RegisterUserGroupList"
            richEdit.ReplaceService(Of IUserGroupListService)(New MyGroupListService())
#End Region ' #RegisterUserGroupList

            richEdit.CreateNewDocument()
            AuthenticateUser()
            CreateRangePermissions()
        End Sub


        Private Sub CreateRangePermissions()
            ' Create document ranges.
            Dim rangeAdmin As DocumentRange = AppendDocument("Documents\administrator.docx")
            Dim rangeBody As DocumentRange = AppendDocument("Documents\body.docx")
            Dim rangeSignature As DocumentRange = AppendDocument("Documents\signature.docx")

            ' Protect document ranges.
#Region "#CreateRangePermissions"
            Dim rangePermissions As RangePermissionCollection = richEdit.Document.BeginUpdateRangePermissions()

            Dim permission As RangePermission = rangePermissions.CreateRangePermission(rangeAdmin)
            permission.UserName = "Nancy Skywalker"
            permission.Group = "Skywalkers"
            rangePermissions.Add(permission)

            Dim permission2 As RangePermission = rangePermissions.CreateRangePermission(rangeBody)
            permission2.Group = "Everyone"
            rangePermissions.Add(permission2)

            Dim permission3 As RangePermission = rangePermissions.CreateRangePermission(rangeSignature)
            permission3.Group = "Nihlus"
            rangePermissions.Add(permission3)

            richEdit.Document.EndUpdateRangePermissions(rangePermissions)
            ' Enforce protection and set password.
            richEdit.Document.Protect("123")
#End Region ' #CreateRangePermissions
        End Sub
        Private Function AppendDocument(ByVal filename As String) As DocumentRange
            richEdit.Document.Paragraphs.Insert(richEdit.Document.Range.End)
            Dim pos As DocumentPosition = richEdit.Document.CreatePosition(richEdit.Document.Range.End.ToInt() - 2)
            Dim range As DocumentRange = richEdit.Document.InsertDocumentContent(pos, filename, DocumentFormat.OpenXml)
            Return range
        End Function
        Private Sub AuthenticateUser()
#Region "#Authentication"
            'Define the user credentials:
            richEdit.AuthenticationOptions.UserName = "Nancy Skywalker"
            richEdit.AuthenticationOptions.Group = "Skywalkers"
#End Region ' #Authentication

#Region "#RangesColor"
            'Customize the editable ranges appearance: 
            richEdit.RangePermissionOptions.HighlightColor = Color.FromArgb(100, 213, 239, 255)
            richEdit.RangePermissionOptions.HighlightBracketsColor = Color.FromArgb(100, 0, 128, 128)
#End Region ' #RangesColor
        End Sub

    End Class
#Region "#NewUserGroupList"
    Friend Class MyGroupListService
        Implements IUserGroupListService

        Private userGroups As List(Of String) = CreateUserGroups()

        Private Shared Function CreateUserGroups() As List(Of String)
            Dim result As New List(Of String)()
            result.Add("Everyone")
            result.Add("Administrators")
            result.Add("Contributors")
            result.Add("Owners")
            result.Add("Editors")
            result.Add("Current User")
            result.Add("Skywalkers")
            result.Add("Nihlus")
            Return result
        End Function
        Public Function GetUserGroups() As IList(Of String)
            Return userGroups
        End Function

        Private Function IUserGroupListService_GetUserGroups() As IList(Of String) Implements IUserGroupListService.GetUserGroups
            Throw New NotImplementedException()
        End Function
    End Class
#End Region ' #NewUserGroupList

#Region "#NewUserList"
    Friend Class MyUserListService
        Implements IUserListService

        Private users As List(Of String) = CreateUsers()

        Private Shared Function CreateUsers() As List(Of String)
            Dim result As New List(Of String)()
            result.Add("Nancy Skywalker")
            result.Add("Andrew Nihlus")
            result.Add("Janet Skywalker")
            result.Add("Margaret")
            Return result
        End Function
        Public Function GetUsers() As IList(Of String)
            Return users
        End Function

        Private Function IUserListService_GetUsers() As IList(Of String) Implements IUserListService.GetUsers
            Throw New NotImplementedException()
        End Function
    End Class
#End Region ' #NewUserList


End Namespace
