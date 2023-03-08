using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace ConsoleCSOM.UserPermissions
{
    public class CsomUserPermissionHelper
    {
        public static async Task AssignUsersToRolesInList(ClientContext ctx, string listTitle, string[] emails, RoleType roleType = RoleType.WebDesigner)
        {
            try
            {
                RoleDefinitionCollection roleDefs = ctx.Web.RoleDefinitions;
                ctx.Load(ctx.Web.RoleDefinitions);
                await ctx.ExecuteQueryAsync();

                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                RoleDefinition roleDef = roleDefs.GetByType(roleType);
                ctx.Load(roleDef);
                await ctx.ExecuteQueryAsync();
                await BreakListInheritance(ctx, targetList);
                foreach (var email in emails)
                {
                    try
                    {
                        var role = new RoleDefinitionBindingCollection(ctx);
                        ctx.Load(role);
                        await ctx.ExecuteQueryAsync();
                        role.Add(roleDef);
                        User user = ctx.Web.EnsureUser(email);
                        targetList.RoleAssignments.Add(user, role);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error when assign role to user {email}");
                        Console.WriteLine(ex.Message);
                    }
                }
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Assign role to users successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task BreakListInheritance(ClientContext ctx, List list)
        {
            try
            {
                list.BreakRoleInheritance(true, false);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Break inheritance in list {list.Title} successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task BreakSiteInheritance(ClientContext ctx)
        {
            try
            {
                ctx.Load(ctx.Web, w => w.HasUniqueRoleAssignments);
                await ctx.ExecuteQueryAsync();
                if (!ctx.Web.HasUniqueRoleAssignments)
                {
                    ctx.Web.BreakRoleInheritance(true, false);
                    await ctx.ExecuteQueryAsync();
                    Console.WriteLine($"Break inheritance in site {ctx.Web.Url} successfully");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task EnableListInheritance(ClientContext ctx, string listTitle)
        {
            try
            {
                List list = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(list, l => l.HasUniqueRoleAssignments, l => l.Title);
                await ctx.ExecuteQueryAsync();

                await EnableListInheritance(ctx, list);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task EnableListInheritance(ClientContext ctx, List list)
        {
            try
            {
                if (list.HasUniqueRoleAssignments)
                {
                    list.ResetRoleInheritance();
                    await ctx.ExecuteQueryAsync();
                }

                Console.WriteLine(
                    $"Delete user inheritance (unique permissions) in list {list.Title} successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreatePermissionLevel(ClientContext ctx, string permissionLevelName,
            List<PermissionKind> permissionKinds, string description = "")
        {
            try
            {
                RoleDefinitionCollection roleDefs = ctx.Site.RootWeb.RoleDefinitions;
                ctx.Load(roleDefs);
                await ctx.ExecuteQueryAsync();

                RoleDefinitionCreationInformation roleDefinitionCreationInformation = new RoleDefinitionCreationInformation();
                roleDefinitionCreationInformation.Name = permissionLevelName;
                roleDefinitionCreationInformation.Description = description;
                roleDefinitionCreationInformation.BasePermissions = new BasePermissions();
                foreach (PermissionKind permissionKind in permissionKinds)
                {
                    roleDefinitionCreationInformation.BasePermissions.Set(permissionKind);
                }

                roleDefs.Add(roleDefinitionCreationInformation);
                await ctx.ExecuteQueryAsync();
                Console.Write($"Create permission level {permissionLevelName} with permissions: ");
                permissionKinds.ForEach(p => Console.Write($"{p.ToString()} "));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateGroup(ClientContext ctx, string groupTitle, string description = "")
        {
            try
            {
                GroupCollection groupCollection = ctx.Web.SiteGroups;

                GroupCreationInformation groupCreationInformation = new GroupCreationInformation
                {
                    Title = groupTitle,
                    Description = description
                };

                Group group = groupCollection.Add(groupCreationInformation);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Create group {groupTitle} successfully");
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }


        public static async Task AssignPermissionLevelToGroups(ClientContext ctx, string groupTitle,
            string permissionLevelName)
        {
            try
            {
                await BreakSiteInheritance(ctx);
                GroupCollection groupCollection = ctx.Web.SiteGroups;
                ctx.Load(groupCollection);
                await ctx.ExecuteQueryAsync();

                Group targetGroup = groupCollection.GetByName(groupTitle);
                ctx.Load(targetGroup);
                await ctx.ExecuteQueryAsync();

                RoleDefinitionCollection roleDefs = ctx.Web.RoleDefinitions;
                ctx.Load(ctx.Web.RoleDefinitions);
                await ctx.ExecuteQueryAsync();

                RoleDefinition roleDef = roleDefs.GetByName(permissionLevelName);

                RoleDefinitionBindingCollection roleDefinitionBindingCollection =
                    new RoleDefinitionBindingCollection(ctx);
                ctx.Load(roleDefinitionBindingCollection);
                await ctx.ExecuteQueryAsync();
                roleDefinitionBindingCollection.Add(roleDef);

                ctx.Web.RoleAssignments.Add(targetGroup, roleDefinitionBindingCollection);

                await ctx.ExecuteQueryAsync();

                Console.WriteLine(
                    $"Assign permission {permissionLevelName} to group {groupTitle} successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task AddUsersToGroup(ClientContext ctx, string groupTitle, List<string> emails)
        {
            try
            {
                GroupCollection groupCollection = ctx.Web.SiteGroups;
                ctx.Load(groupCollection);
                await ctx.ExecuteQueryAsync();

                Group targetGroup = groupCollection.GetByName(groupTitle);
                ctx.Load(targetGroup);
                await ctx.ExecuteQueryAsync();

                foreach (string email in emails)
                {
                    try
                    {
                        User user = ctx.Web.EnsureUser(email);
                        targetGroup.Users.AddUser(user);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                }

                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Add users to group {groupTitle} successfully");
                emails.ForEach(email => Console.Write($" {email} "));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

    }
}
