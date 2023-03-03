using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ContentType = Microsoft.SharePoint.Client.ContentType;

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

        public static async Task CreateList(ClientContext ctx, string title, string description = "")
        {
            try
            {
                var listCreationInfo = new ListCreationInformation
                {
                    Title = title,
                    TemplateType = (int)ListTemplateType.GenericList,
                    Description = description,
                };

                var list = ctx.Web.Lists.Add(listCreationInfo);
                list.ContentTypesEnabled = true;
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

        public static async Task CreateTermSet(ClientContext ctx, string termSetName)
        {
            try
            {
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

        public static async Task CreateCityTerm(ClientContext ctx, string termSetName, string cityName)
        {
            try
            {
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

        public static async Task CreateSiteFields(ClientContext ctx, FieldType fieldType, string fieldName)
        {
            try
            {
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
        public static async Task CreateTaxonomySiteField(ClientContext ctx, string fieldName,
            string termSetName)
        {
            try
            {
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

        public static async Task CreateContentTypeForList(ClientContext ctx, string name, string group = "Custom Content Types", string description = "")
        {
            try
            {
                ContentTypeCollection contentTypes = ctx.Web.ContentTypes;
                ctx.Load(contentTypes);
                await ctx.ExecuteQueryAsync();

                // README: To know more about content type ID: https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ms452896(v=office.14)
                var itemContentTypeId = contentTypes.GetById("0x01");

                ctx.Load(itemContentTypeId);
                await ctx.ExecuteQueryAsync();

                var newContentType = new ContentTypeCreationInformation()
                {
                    Name = name,
                    ParentContentType = itemContentTypeId,
                    Group = group,
                    Description = description
                };

                var addContentType = ctx.Site.RootWeb.ContentTypes.Add(newContentType);
                ctx.Load(addContentType);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Content type {name} created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static async Task AddFieldsToContentTypeByName(ClientContext ctx, string contentTypeName, string fieldName)
        {
            try
            {
                ContentTypeCollection contentTypeCollection = ctx.Web.ContentTypes;
                ctx.Load(contentTypeCollection);
                await ctx.ExecuteQueryAsync();

                var targetContentType = contentTypeCollection.Single(ct => ct.Name == contentTypeName);

                ctx.Load(targetContentType);
                await ctx.ExecuteQueryAsync();

                Field targetField = ctx.Web.AvailableFields.GetByInternalNameOrTitle(fieldName);
                FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                fldLink.Field = targetField;
                targetContentType.FieldLinks.Add(fldLink);
                targetContentType.Update(false);

                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Field {fieldName} added to content type {contentTypeName} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        public static async Task MakeContentTypeDefaultInList(ClientContext ctx, string listName, string contentTypeName)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listName);
                var currentContentTypeOrder = targetList.ContentTypes;
                ctx.Load(currentContentTypeOrder, coll => coll.Include(
                    ct => ct.Name,
                    ct => ct.Id));
                await ctx.ExecuteQueryAsync();

                IList<ContentTypeId> reverseOrder = (
                    from ct in currentContentTypeOrder 
                    where ct.Name.Equals(contentTypeName, StringComparison.OrdinalIgnoreCase)
                    select ct.Id).ToList();
                targetList.RootFolder.UniqueContentTypeOrder = reverseOrder;
                targetList.RootFolder.Update();
                targetList.Update();
                await ctx.ExecuteQueryAsync();
                Console.WriteLine(
                    $"Make Content Type {contentTypeName} default in list {listName} successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        public static async Task AddContentTypeToListByName(ClientContext ctx, string listName, string contentTypeName)
        {
            // TODO: Add content type to specific lists
            try
            {
                ContentTypeCollection contentTypeCollection = ctx.Web.ContentTypes;
                ctx.Load(contentTypeCollection);
                await ctx.ExecuteQueryAsync();

                ContentType targetContentType = contentTypeCollection.Single(ct => ct.Name == contentTypeName);
                List targetList = ctx.Web.Lists.GetByTitle(listName);

                targetList.ContentTypes.AddExistingContentType(targetContentType);
                targetList.Update();
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Content Type {contentTypeName} added to list {listName} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }


        public static async Task Init5ItemsToList(ClientContext ctx)
        {
            // TODO: Add 5 items to the list above

        }

        public static async Task UpdateAboutFieldDefaultValue(ClientContext ctx)
        {
            // TODO: Update default value for about site fields
        }

        public static async Task UpdateCityFieldDefaultValue(ClientContext ctx)
        {
            // TODO: Update city value for city site fields
        }
    }
}
