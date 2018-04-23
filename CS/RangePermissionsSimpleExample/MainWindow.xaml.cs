using System.Collections.Generic;
using DevExpress.Xpf.Ribbon;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Services;
using System.Windows.Media;
namespace RangePermissionsSimpleExample
{
    public partial class MainWindow : DXRibbonWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            #region #RegisterUserList
            richEdit.ReplaceService<IUserListService>(new MyUserListService());
            #endregion #RegisterUserList

            #region #RegisterUserGroupList
            richEdit.ReplaceService<IUserGroupListService>(new MyGroupListService());
            #endregion #RegisterUserGroupList

            richEdit.CreateNewDocument();
            AuthenticateUser();
            CreateRangePermissions();
        }


        private void CreateRangePermissions()
        {
            // Create document ranges.
            DocumentRange rangeAdmin = AppendDocument("Documents\\administrator.docx");
            DocumentRange rangeBody = AppendDocument("Documents\\body.docx");
            DocumentRange rangeSignature = AppendDocument("Documents\\signature.docx");

            // Protect document ranges.
            #region #CreateRangePermissions            
            RangePermissionCollection rangePermissions = richEdit.Document.BeginUpdateRangePermissions();

            RangePermission permission = rangePermissions.CreateRangePermission(rangeAdmin);
            permission.UserName = "Nancy Skywalker";
            permission.Group = "Skywalkers";
            rangePermissions.Add(permission);

            RangePermission permission2 = rangePermissions.CreateRangePermission(rangeBody);
            permission2.Group = @"Everyone";
            rangePermissions.Add(permission2);

            RangePermission permission3 = rangePermissions.CreateRangePermission(rangeSignature);
            permission3.Group = "Nihlus";
            rangePermissions.Add(permission3);

            richEdit.Document.EndUpdateRangePermissions(rangePermissions);
            // Enforce protection and set password.
            richEdit.Document.Protect("123");
            #endregion #CreateRangePermissions
        }
        private DocumentRange AppendDocument(string filename)
        {
            richEdit.Document.Paragraphs.Insert(richEdit.Document.Range.End);
            DocumentPosition pos = richEdit.Document.CreatePosition(richEdit.Document.Range.End.ToInt() - 2);
            DocumentRange range = richEdit.Document.InsertDocumentContent(pos, filename, DocumentFormat.OpenXml);
            return range;
        }
        private void AuthenticateUser()
        {
            #region #Authentication
            //Define the user credentials:
            richEdit.AuthenticationOptions.UserName = "Nancy Skywalker";
            richEdit.AuthenticationOptions.Group = "Skywalkers";
            #endregion #Authentication

            #region #RangesColor
            //Customize the editable ranges appearance: 
            richEdit.RangePermissionOptions.HighlightColor = Color.FromArgb(100, 213, 239, 255);
            richEdit.RangePermissionOptions.HighlightBracketsColor = Color.FromArgb(100, 0, 128, 128);
            #endregion #RangesColor
        }

    }
    #region #NewUserGroupList
    class MyGroupListService : IUserGroupListService
    {
        List<string> userGroups = CreateUserGroups();

        static List<string> CreateUserGroups()
        {
            List<string> result = new List<string>();
            result.Add(@"Everyone");
            result.Add(@"Administrators");
            result.Add(@"Contributors");
            result.Add(@"Owners");
            result.Add(@"Editors");
            result.Add(@"Current User");
            result.Add("Skywalkers");
            result.Add("Nihlus");
            return result;
        }
        public IList<string> GetUserGroups()
        {
            return userGroups;
        }
    }
    #endregion #NewUserGroupList

    #region #NewUserList
    class MyUserListService : IUserListService
    {
        List<string> users = CreateUsers();

        static List<string> CreateUsers()
        {
            List<string> result = new List<string>();
            result.Add("Nancy Skywalker");
            result.Add("Andrew Nihlus");
            result.Add("Janet Skywalker");
            result.Add("Margaret");
            return result;
        }
        public IList<string> GetUsers()
        {
            return users;
        }
    }
    #endregion #NewUserList


}
