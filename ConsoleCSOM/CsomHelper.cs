using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public class CsomHelper
    {
        private static readonly int Lcid = 1033;
        public static TermSet GetTermSet(ClientContext ctx, string termSetName)
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
                //Console.Error.WriteLine(ex.Message);
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
            try
            {
                var termSet = GetTermSet(ctx, termSetName);
                var term = termSet.Terms.GetByName(termName);
                ctx.Load(term);
                ctx.ExecuteQuery();
                return term != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public static bool CheckSiteFieldNameExists(ClientContext ctx, string siteFieldName)
        {
            try
            {
                var fieldExists = ctx.Web.Fields.GetByInternalNameOrTitle(siteFieldName);
                ctx.Load(fieldExists);
                ctx.ExecuteQuery();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
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

                Console.WriteLine($"List {title} created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateTermSetCsom(ClientContext ctx, string termSetName)
        {
            try
            {
                if (CheckTermSetNameExists(ctx, termSetName))
                {
                    Console.WriteLine($"Termset {termSetName} already exists!");
                    return;
                }

                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                ctx.Load(taxonomySession);
                await ctx.ExecuteQueryAsync();

                if (taxonomySession != null)
                {
                    TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();

                    if (termStore != null)
                    {
                        TermGroup termGroup = termStore.GetSiteCollectionGroup(ctx.Site, true);
                        termGroup.CreateTermSet(termSetName, Guid.NewGuid(), Lcid); // Lcid: English
                        await ctx.ExecuteQueryAsync();
                    }
                }
                Console.WriteLine($"Termset {termSetName} created successfully!");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateCityTermCsom(ClientContext ctx, string termSetName, string cityName)
        {
            try
            {
                if (CheckTermExistsInTermSetWithName(ctx, termSetName, cityName))
                {
                    Console.WriteLine($"Term {cityName} already exists in term set {termSetName}");
                    return;
                }

                var termSet = GetTermSet(ctx, termSetName);
                Guid guid = Guid.NewGuid();
                var term = termSet.CreateTerm(cityName, Lcid, guid);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Term {cityName} in set {termSetName} created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateSiteFieldsCsom(ClientContext ctx, FieldType fieldType, string fieldName)
        {
            try
            {
                if (CheckSiteFieldNameExists(ctx, fieldName))
                {
                    Console.WriteLine($"Field {fieldName} already exists");
                    return;
                }

                var createField = ctx.Web.Fields.AddFieldAsXml(
                    $"<Field Type='{fieldType.ToString()}' DisplayName='{fieldName}' Name='{fieldName}' />",
                    true, AddFieldOptions.DefaultValue);
                ctx.Load(createField);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Field {fieldName} created");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // create taxonomy site field
        public static async Task CreateTaxonomySiteFieldCsom(ClientContext ctx, string fieldName,
            string termSetName)
        {
            try
            {
                if (CheckSiteFieldNameExists(ctx, fieldName))
                {
                    Console.WriteLine($"Field {fieldName} already exists");
                    return;
                }

                var termSet = GetTermSet(ctx, termSetName);
                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                ctx.Load(termStore);
                await ctx.ExecuteQueryAsync();

                var createField = ctx.Web.Fields.AddFieldAsXml(
                    $"<Field Type='TaxonomyFieldType' DisplayName='{fieldName}' Name='{fieldName}' StaticName='{fieldName}' TermSetId='{termSet.Id.ToString()}' />",
                    true, AddFieldOptions.DefaultValue);

                ctx.Load(createField);
                await ctx.ExecuteQueryAsync();

                var updateTaxField = ctx.CastTo<TaxonomyField>(createField);
                updateTaxField.SspId = termStore.Id;
                updateTaxField.TermSetId = termSet.Id;
                updateTaxField.Update();
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Taxonomy Field {fieldName} created!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
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
