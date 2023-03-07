using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;


namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        private static readonly string  ContentTypeNameDefault = "CSOM Test Content Type";
        private static readonly string  DocumentListName = "Document Test";
        private static readonly string Folder2Url = "/Folder 1/Folder 2";
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);

                    User currentUser = ctx.Web.CurrentUser;
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    ctx.Load(currentUser);
                    await ctx.ExecuteQueryAsync();
                    // Create list
                    //await CsomHelper.CreateList(ctx, $"CSOM Test",
                    //    "Practice CSOM create list");

                    // Create term set
                    string termSetName = $"city-{currentUser.Title}";
                    //await CsomHelper.CreateTermSet(ctx, termSetName);
                    //await CsomHelper.CreateCityTerm(ctx, termSetName, "Ho Chi Minh");
                    //await CsomHelper.CreateCityTerm(ctx, termSetName, "Stockholm");

                    ////create site fields
                    ////await CsomHelper.CreateSiteFields(ctx, FieldType.Text, "about");
                    //await CsomHelper.CreateTaxonomySiteField(ctx, "city", termSetName);

                    //// create content type
                    //await CsomHelper.CreateContentTypeForList(ctx, ContentTypeNameDefault);

                    //// add field site to content type
                    //await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "city");
                    //await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "about");

                    //// add content type to list
                    //await CsomHelper.AddContentTypeToListByName(ctx, "CSOM Test", "CSOM Test Content Type");
                    //await CsomHelper.MakeContentTypeDefaultInList(ctx, "CSOM Test", "CSOM Test Content Type");

                    //await CsomHelper.InitItemsToList(ctx, "CSOM Test", termSetName, "Stockholm", "", 10);

                    //await CsomHelper.UpdateFieldDefaultValue(ctx, "city", "Ho Chi Minh");

                    //await CsomHelper.QueryNotAboutItemsCaml(ctx, "CSOM Test");

                    //await CsomHelper.CreateListViewByCityOrderByCreatedTime(ctx, "CSOM Test",
                    //    "Ho Chi Minh View ###12", "Ho Chi Minh");

                    //await CsomHelper.UpdateBatchAboutColumnCaml(ctx, "CSOM Test", "Update Script");

                    //await CsomHelper.CreateTaxonomyFieldMulti(ctx, "cities", termSetName);

                    //await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "cities");

                    //await CsomHelper.InitItemsToList(ctx, "CSOM Test", termSetName, "", "", 2, true, "Ho Chi Minh,Stockholm1");

                    //await CsomHelper.CreateDocumentLibraryListType(ctx, DocumentListName);

                    //await CsomHelper.AddContentTypeToListByName(ctx, DocumentListName, "CSOM Test Content Type");

                    //await CsomHelper.CreateFolderInList(ctx, DocumentListName, "Folder 2", "/Folder 1");
                    //await CsomHelper.CreateFolderInList(ctx, DocumentListName, "Folder 1.3");

                    
                    // create 3 items
                    //for (int i = 0; i < 3; i++)
                    //{
                    //    await CsomHelper.CreateItemInFolder(ctx, DocumentListName, Folder2Url, $"hello {i} plus", "Folder Test");
                    //}

                    //for (int i = 0; i < 2; i++)
                    //{
                    //    await CsomHelper.CreateItemInFolder(ctx, DocumentListName, Folder2Url, $"hello taxonomy {i}", "Folder Test", "Stockholm", termSetName);
                    //}

                    //await CsomHelper.GetItemsWithStockHolmInFolderTask(ctx, Folder2Url, DocumentListName);

                    //await CsomHelper.CreateItemInFolder(ctx, DocumentListName, "", "This is at root");
                }

                Console.WriteLine("Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }

        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
    }
}
