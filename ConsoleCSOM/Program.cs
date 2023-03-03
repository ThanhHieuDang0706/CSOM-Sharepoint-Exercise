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

                    // create site fields
                    //await CsomHelper.CreateSiteFields(ctx, FieldType.Text, "about", termSetName);
                    //await CsomHelper.CreateTaxonomySiteField(ctx, "city", termSetName);

                    // create content type
                    //await CsomHelper.CreateContentTypeForList(ctx, ContentTypeNameDefault);

                    //// add field site to content type
                    //await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "city");
                    //await CsomHelper.AddFieldsToContentTypeByName(ctx, ContentTypeNameDefault, "about");

                    // add content type to list
                    await CsomHelper.AddContentTypeToListByName(ctx, "CSOM Test", "CSOM Test Content Type");
                    await CsomHelper.MakeContentTypeDefaultInList(ctx, "CSOM Test", "CSOM Test Content Type");
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

        //private static async Task CsomLinqAsync(ClientContext ctx)
        //{
        //    var fieldsQuery = from f in ctx.Web.Fields
        //                      where f.InternalName == "Test" ||
        //                            f.TypeAsString == "TaxonomyFieldTypeMulti" ||
        //                            f.TypeAsString == "TaxonomyFieldType"
        //                      select f;

        //    var fields = ctx.LoadQuery(fieldsQuery);
        //    await ctx.ExecuteQueryAsync();
        //}

        //private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        //{
        //    var list = ctx.Web.Lists.GetByTitle("Documents");

        //    var allItemsQuery = CamlQuery.CreateAllItemsQuery();
        //    var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

        //    var items = list.GetItems(new CamlQuery()
        //    {
        //        ViewXml = @"<View>
        //                        <Query>
        //                            <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
        //                        </Query>
        //                        <RowLimit>20</RowLimit>
        //                    </View>",
        //        FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
        //        //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
        //    });

        //    ctx.Load(items);
        //    await ctx.ExecuteQueryAsync();
        //}



    }
}
