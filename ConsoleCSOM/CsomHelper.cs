using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    internal class CsomHelper
    {
        public static async Task CreateListCsom(ClientContext ctx, string title, string description = "")
        {
            // TODO: Create list with title and description
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

        public static async Task CreateTermSetCsom(ClientContext ctx)
        {
            // TODO: Create term set
            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
            Guid newTermSetId = Guid.NewGuid();
            group.CreateTermSet($"city-{ctx.Web.CurrentUser.Title}", newTermSetId, newTermSetId.GetHashCode());
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateCityTermCsom(ClientContext ctx, string cityName)
        {
            // TODO: Create city term
            var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            var group = termStore.GetSiteCollectionGroup(ctx.Site, true);
            var termSet = group.TermSets.GetByName($"city-{ctx.Web.CurrentUser.Title}");
            Guid guid = Guid.NewGuid();
            var term = termSet.CreateTerm(cityName, guid.GetHashCode(), guid);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateSiteFieldsCsom(ClientContext ctx)
        {
            // TODO: Create Site Fields
            

        }

        public static async Task CreateContentTypeCsom(ClientContext ctx)
        {
            // TODO: Create Content Type
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
