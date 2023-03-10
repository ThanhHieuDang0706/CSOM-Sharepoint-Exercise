using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ConsoleCSOM.UserPermissions;


namespace ConsoleCSOM
{
    class CsomUserPermissionExerciseRunner
    {
        private static string[] DesignerEmails = GetEmailsOfUsersToAssignRole("EmailToHaveDesignPermission");
        private static List<string> GroupEmails = GetEmailsOfUsersToAssignRole("EmailToBeAddedInGroup").ToList();
        private static string ListTitle = "Accounts";
        private static string PermissionLevelName = "Test Level";
        private static string GroupTitle = "First Create";
        private static List<PermissionKind> TestPermissionKinds = new List<PermissionKind>()
        {
            PermissionKind.ManageLists,
            PermissionKind.CreateAlerts
        };
        public static async Task Run()
        {
            using (var clientContextHelper = new ClientContextHelper())
            {
                ClientContext ctx = Program.GetContext(clientContextHelper, "SharepointInfoUserPermissionExercise");

                User currentUser = ctx.Web.CurrentUser;
                ctx.Load(ctx.Web);
                await ctx.ExecuteQueryAsync();

                ctx.Load(currentUser);
                await ctx.ExecuteQueryAsync();

                await RunTask(currentUser, ctx);
            }
        }

        private static async Task RunTask(User currentUser, ClientContext ctx)
        {
            //await CsomUserPermissionHelper.AssignUsersToRolesInList(ctx, ListTitle, DesignerEmails);
            //await CsomUserPermissionHelper.EnableListInheritance(ctx, ListTitle);
            //await CsomUserPermissionHelper.CreatePermissionLevel(ctx, PermissionLevelName,
            //    TestPermissionKinds);

            //await CsomUserPermissionHelper.CreateGroup(ctx, GroupTitle);
            //await CsomUserPermissionHelper.AssignPermissionLevelToGroups(ctx, GroupTitle, PermissionLevelName);

            //await CsomUserPermissionHelper.AddUsersToGroup(ctx, GroupTitle, GroupEmails);
        }

        private static string[] GetEmailsOfUsersToAssignRole(string key)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var emails = config.GetSection(key).Get<string[]>();
            return emails;
        }
    }
}
