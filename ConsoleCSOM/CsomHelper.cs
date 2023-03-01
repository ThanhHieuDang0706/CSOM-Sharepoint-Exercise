using Microsoft.SharePoint.Client;
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
        }

        public static async Task CreateSiteFieldsCsom(ClientContext ctx)
        {
            // TODO: Create Site Fields
            //var field = ctx.Web.Fields.AddFieldAsXml("<Field Type='TaxonomyFieldType' DisplayName='MyTaxonomyField' Required='FALSE' EnforceUniqueValues='FALSE' List='MyList' ShowField='Term1033' Mult='TRUE' />", true, AddFieldOptions.DefaultValue);
            //ctx.Load(field);
            //await ctx.ExecuteQueryAsync();
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

        public static async Task UpdateSiteFieldDefaultValueCsom(ClientContext ctx)
        {
            // TODO: Update default value for site fields
        }
    }
}
