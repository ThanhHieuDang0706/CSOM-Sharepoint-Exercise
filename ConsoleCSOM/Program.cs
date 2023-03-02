using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using System.Runtime.CompilerServices;
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
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = GetContext(clientContextHelper);

                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();
                    User currentUser = ctx.Web.CurrentUser;

                    ctx.Load(currentUser);
                    await ctx.ExecuteQueryAsync();

                    // Create list
                    //await CreateListCsom(ctx, $"CSOM Test",
                    //    "Practice CSOM create list");

                    // Create term set
                    await CreateTermSetCsom(ctx, $"city-{currentUser.Title}");


                }

                Console.WriteLine($"Press Any Key To Stop!");
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



        //private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        //{
        //    //load term by id
        //    TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
        //    Term taxonomyTerm = session.GetTerm(new Guid(termId));
        //    Ctx.Load(taxonomyTerm, t => t.Labels,
        //                           t => t.Name,
        //                           t => t.Id);
        //    await Ctx.ExecuteQueryAsync();
        //}

        //private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        //{
        //    var field = ctx.Web.Fields.GetByTitle("fieldname");

        //    ctx.Load(field);
        //    await ctx.ExecuteQueryAsync();

        //    var taxField = ctx.CastTo<TaxonomyField>(field);

        //    taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
        //    {
        //        WssId = -1, // alway let it -1
        //        Label = "correct label here",
        //        TermGuid = "term id"
        //    });
        //    item.Update();
        //    await ctx.ExecuteQueryAsync();
        //}

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");

            var terms = termSet.GetAllTerms();
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
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

        public static TermSet getTermSet(ClientContext ctx, string termSetName)
        {
            try
            {
                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
                var termSet = group.TermSets.GetByName(termSetName);
                ctx.Load(termSet);
                ctx.ExecuteQuery();
                return termSet;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return null;
            }
        }

        public static bool CheckListNameExists(ClientContext ctx, string listTitle)
        {
            try
            {
                var list = ctx.Web.Lists.Single(l => l.Title == listTitle);
                ctx.Load(list);
                ctx.ExecuteQuery();
                return list != null;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return false;
            }
        }

        public static bool CheckTermSetNameExists(ClientContext ctx, string termSetName)
        {
            try
            {
                // get site 
                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
                var termSet = group.TermSets.GetByName(termSetName);
                ctx.Load(termSet);
                ctx.ExecuteQuery();
                return termSet != null;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return false;
            }
        }

        public static bool CheckTermExistsInTermSetWithName(ClientContext ctx, string termSetName,
            string termName)
        {
            var termSet = getTermSet(ctx, termSetName);
            var term = termSet.Terms.GetByName(termName);
            ctx.Load(term);
            ctx.ExecuteQuery();
            return term != null;
        }

        public static bool CheckContentTypeExists(ClientContext ctx, string contentTypeName)
        {
            var contentType = ctx.Web.ContentTypes.Single(ct => ct.Name == contentTypeName);
            ctx.Load(contentType);
            ctx.ExecuteQuery();
            return contentType != null;
        }

        public static async Task CreateListCsom(ClientContext ctx, string title, string description = "")
        {
            // TODO: Create list with title and description
            try
            {
                if (CheckListNameExists(ctx, title))
                {
                    Console.WriteLine($"List {title} already exists");
                    return;
                }

                var listCreationInfo = new ListCreationInformation
                {
                    Title = title,
                    TemplateType = (int)ListTemplateType.GenericList,
                    Description = description
                };

                var list = ctx.Web.Lists.Add(listCreationInfo);

                list.Description = description;
                list.Update();
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateTermSetCsom(ClientContext ctx, string termSetName)
        {
            // TODO: Create term set
            try
            {
                if (CheckTermSetNameExists(ctx, termSetName))
                {
                    Console.WriteLine($"Term set {termSetName} already exists");
                    return;
                }

                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
                Guid newTermSetId = Guid.NewGuid();

                group.CreateTermSet(termSetName, newTermSetId, 1000);
                await ctx.ExecuteQueryAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateCityTermCsom(ClientContext ctx, string termSetName, string cityName)
        {
            // TODO: Create city term
            if (CheckTermExistsInTermSetWithName(ctx, termSetName, cityName))
            {
                Console.WriteLine($"Term {cityName} already exists");
                return;
            }

            var termSet = getTermSet(ctx, termSetName);
            Guid guid = Guid.NewGuid();
            var term = termSet.CreateTerm(cityName, guid.GetHashCode(), guid);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateSiteFieldsCsom(ClientContext ctx, FieldType fieldType, string fieldName)
        {
            // TODO: Create Site Fields
            // Check if the name field exists
            var fieldExists = ctx.Web.Fields.GetByInternalNameOrTitle(fieldName);
            ctx.Load(fieldExists);
            await ctx.ExecuteQueryAsync();
            if (fieldExists != null)
            {
                Console.WriteLine($"Field {fieldName} already exists");
                return;
            }

            var createField = ctx.Web.Fields.AddFieldAsXml($"<Field Type='{fieldType.ToString()}' DisplayName='{fieldName}' Name='{fieldName}' />", true, AddFieldOptions.DefaultValue);
            ctx.Load(createField);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateContentTypeCsom(ClientContext ctx, string name, string group = "Custom Content Types", string description = "")
        {
            // TODO: Create Content Type
            if (CheckContentTypeExists(ctx, name))
            {
                Console.WriteLine($"Content type {name} already exists");
                return;
            }

            // README: To know more about content type ID: https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ms452896(v=office.14)
            var itemContentTypeId = ctx.Web.AvailableContentTypes.GetById("0x01");

            var newContentType = new ContentTypeCreationInformation()
            {
                Name = name,
                Id = itemContentTypeId.StringId,
                Group = group,
                Description = description
            };

            var addContentType = ctx.Web.ContentTypes.Add(newContentType);
            ctx.Load(addContentType);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddContentTypeToListCsom(ClientContext ctx)
        {
            // TODO: Add content type to specific lists

        }

        public static async Task AddFieldsToContentTypeCsom(ClientContext ctx)
        {
            // TODO: Add site fields to content type
        }

        public static async Task Init5ItemsToList(ClientContext ctx)
        {
            // TODO: Add 5 items to the list above

        }

        public static async Task UpdateAboutFieldDefaultValueCsom(ClientContext ctx)
        {
            // TODO: Update default value for about site fields
        }

        public static async Task UpdateCityFieldDefaultValueCsom(ClientContext ctx)
        {
            // TODO: Update city value for city site fields
        }

    }
}
