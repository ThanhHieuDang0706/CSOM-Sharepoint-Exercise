using ConsoleCSOM;
using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

namespace ConsoleCSOM.Csom
{
    class CsomExerciseRunner
    {
        private static readonly string ContentTypeNameDefault = "CSOM Test Content Type";
        private static readonly string DocumentListName = "Document Test";
        private static readonly string Folder2Url = "/Folder 1/Folder 2";
        private static readonly string ListTestName = "CSOM Test";


        public static async Task Run()
        {
            using (var clientContextHelper = new ClientContextHelper())
            {
                ClientContext ctx = Program.GetContext(clientContextHelper, "SharepointInfo");

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
            // Create list
            await CsomHelper.CreateList(ctx, $"CSOM Test",
                "Practice CSOM create list");

            // Create term set
            string termSetName = $"city-{currentUser.Title}";
            await CsomHelper.CreateTermSet(ctx, termSetName);
            await CsomHelper.CreateCityTerm(ctx, termSetName, "Ho Chi Minh");
            await CsomHelper.CreateCityTerm(ctx, termSetName, "Stockholm");

            //create site fields
            //await CsomHelper.CreateSiteFields(ctx, FieldType.Text, "about");
            await CsomHelper.CreateTaxonomySiteField(ctx, "city", termSetName);

            // create content type
            await CsomHelper.CreateContentTypeForList(ctx, ContentTypeNameDefault);

            // add field site to content type
            await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "city");
            await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "about");

            // add content type to list
            await CsomHelper.AddContentTypeToListByName(ctx, ListTestName, ContentTypeNameDefault);
            await CsomHelper.MakeContentTypeDefaultInList(ctx, ListTestName, ContentTypeNameDefault);

            await CsomHelper.InitItemsToList(ctx, ListTestName, termSetName, "Stockholm", "", 10);

            await CsomHelper.UpdateFieldDefaultValue(ctx, "city", "Ho Chi Minh");

            await CsomHelper.QueryNotAboutItemsCaml(ctx, ListTestName);

            await CsomHelper.CreateListViewByCityOrderByCreatedTime(ctx, ListTestName,
                "Ho Chi Minh View ###12", "Ho Chi Minh");

            await CsomHelper.UpdateBatchAboutColumnCaml(ctx, ListTestName, "Update Script");

            await CsomHelper.CreateTaxonomyFieldMulti(ctx, "cities", termSetName);

            await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "cities");

            await CsomHelper.InitItemsToList(ctx, ListTestName, termSetName, "", "", 2, true, "Ho Chi Minh,Stockholm1");

            await CsomHelper.CreateDocumentLibraryListType(ctx, DocumentListName);

            await CsomHelper.AddContentTypeToListByName(ctx, DocumentListName, ContentTypeNameDefault);

            await CsomHelper.CreateFolderInList(ctx, DocumentListName, "Folder 2", "/Folder 1");
            await CsomHelper.CreateFolderInList(ctx, DocumentListName, "Folder 1.3");


            //create 3 items
            for (int i = 0; i < 3; i++)
            {
                await CsomHelper.CreateItemInFolder(ctx, DocumentListName, Folder2Url, $"hello {i} plus", "Folder Test");
            }

            for (int i = 0; i < 2; i++)
            {
                await CsomHelper.CreateItemInFolder(ctx, DocumentListName, Folder2Url, $"hello taxonomy {i}", "Folder Test", "Stockholm", termSetName);
            }

            await CsomHelper.GetItemsWithStockHolmInFolderTask(ctx, Folder2Url, DocumentListName);

            await CsomHelper.CreateItemInFolder(ctx, DocumentListName, "", "This is at root");

            await CsomHelper.CreateFieldInListCsomHelperTask(ctx, ListTestName, "author",
                FieldType.User, "CSOM Test Author");

            await CsomHelper.GetUserFromEmailOrName(ctx, "Hieu Dang Thanh");

            await CsomHelper.CreateFolderViewAndMakeDefaultView(ctx, DocumentListName, "Folder View 1.2.1");
        }
    }
}
