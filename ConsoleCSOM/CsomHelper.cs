using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    internal class CsomHelper
    {
        public static TermSet getTermSet(ClientContext ctx, string termSetName)
        {
            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
            var termSet = group.TermSets.GetByName(termSetName);
            ctx.Load(termSet);
            ctx.ExecuteQuery();
            return termSet;
        }

        public static bool CheckListNameExists(ClientContext ctx, string listTitle)
        {
            var list = ctx.Web.Lists.GetByTitle(listTitle);
            ctx.Load(list);
            ctx.ExecuteQuery();
            return list != null;
        }

        public static bool CheckTermSetNameExists(ClientContext ctx, string termSetName)
        {
            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
            var termSet = group.TermSets.GetByName(termSetName);
            ctx.Load(termSet);
            ctx.ExecuteQuery();
            return termSet != null;
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
            if (CheckListNameExists(ctx, title))
            {
                Console.WriteLine($"List {title} already exists");
                return;
            }

            var listCreationInfo = new ListCreationInformation
            {
                Title = title,
                TemplateType = (int)ListTemplateType.GenericList
            };

            var list = ctx.Web.Lists.Add(listCreationInfo);
            list.Description = description;
            list.Update();
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateTermSetCsom(ClientContext ctx, string termSetName)
        {
            // TODO: Create term set
            if (CheckTermSetNameExists(ctx, termSetName))
            {
                Console.WriteLine($"Term set {termSetName} already exists");
                return;
            }
            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
            Guid newTermSetId = Guid.NewGuid();
            group.CreateTermSet(termSetName, newTermSetId, newTermSetId.GetHashCode());
            await ctx.ExecuteQueryAsync();
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
