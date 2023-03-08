using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using System.Globalization;


using Task = System.Threading.Tasks.Task;

namespace ConsoleCSOM.Csom
{
    public class CsomHelper
    {
        private static readonly int Lcid = CultureInfo.CurrentCulture.LCID;
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
                Console.WriteLine(ex.Message);
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
                    $"<Field Type='TaxonomyFieldType' DisplayName='{fieldName}' Name='{fieldName}' StaticName='{fieldName}' TermSetId='{{{termSet.Id}}}' />",
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
        public static async Task AddFieldsToContentTypeByName(ClientContext ctx, string contentTypeName, string fieldName, bool updateChildren = true)
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
                targetContentType.Update(updateChildren);

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


        public static async Task InitItemsToList(ClientContext ctx, string listTitle, string termSetName, string city = "", string about = "", int numOfItems = 5, bool addMultiTaxField = false, string commaSeperatedMultiTaxValue = "", string multiTaxFieldName = "cities")
        {
            TaxonomyField taxField =
                ctx.CastTo<TaxonomyField>(ctx.Site.RootWeb.Fields.GetByTitle("city"));
            ctx.Load(taxField);
            await ctx.ExecuteQueryAsync();

            // GET THE TERMS
            TermSet termSet = GetTermSet(ctx, termSetName);
            TermCollection termCollection = termSet.Terms;
            ctx.Load(termCollection);
            await ctx.ExecuteQueryAsync();

            Term defaultTermCityValue = termCollection.GetByName(taxField.DefaultValue);
            ctx.Load(defaultTermCityValue);
            await ctx.ExecuteQueryAsync();

            Term targetTerm;
            try
            {
                targetTerm = termCollection.GetByName(city);
                ctx.Load(targetTerm);
                await ctx.ExecuteQueryAsync();
            }
            catch
            {
                targetTerm = defaultTermCityValue;
            }


            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                for (int i = 0; i < numOfItems; i++)
                {
                    var newCityTaxonomyValue = new TaxonomyFieldValue();
                    if (city == "")
                    {
                        newCityTaxonomyValue.Label = defaultTermCityValue.Name;
                        newCityTaxonomyValue.TermGuid = defaultTermCityValue.Id.ToString();
                        newCityTaxonomyValue.WssId = -1;
                    }

                    else
                    {
                        newCityTaxonomyValue.Label = targetTerm.Name;
                        newCityTaxonomyValue.TermGuid = targetTerm.Id.ToString();
                        newCityTaxonomyValue.WssId = -1;
                    }

                    ListItem newItem = targetList.AddItem(itemCreateInfo);
                    string newGuid = Guid.NewGuid().ToString();

                    if (about != "")
                    {
                        newItem["about"] = about;
                    }
                    else
                    {
                        newItem["about"] = await GetSiteFieldDefaultValue(ctx, "about");
                    }

                    newItem["Title"] = newGuid;
                    newItem["city"] = newCityTaxonomyValue;
                    newItem.Update();
                    if (addMultiTaxField)
                    {
                        // Todo: add multi tax value here: https://stackoverflow.com/questions/17076509/set-taxonomy-field-multiple-values
                        var multiTaxField = ctx.CastTo<TaxonomyField>(targetList.Fields.GetByInternalNameOrTitle("cities"));

                        var listValue = commaSeperatedMultiTaxValue.Split(',');
                        List<string> fieldValueList = new List<string>();
                        foreach (var item in listValue)
                        {
                            try
                            {
                                Term term = termCollection.GetByName(item);
                                ctx.Load(term);
                                await ctx.ExecuteQueryAsync();
                                string fieldValueItem = $"-1;#{term.Name}|{term.Id.ToString()}";
                                fieldValueList.Add(fieldValueItem);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                Console.WriteLine($"Error while loading term {item}!");
                            }
                        }
                        // README: https://sharepoint.stackexchange.com/questions/276504/the-given-value-for-a-taxonomy-field-was-not-formatted-in-the-required-intl
                        string fieldValue = string.Join(";#", fieldValueList);
                        TaxonomyFieldValueCollection newMultiTaxonomyValues =
                            new TaxonomyFieldValueCollection(ctx, fieldValue, multiTaxField);
                        multiTaxField.SetFieldValueByValueCollection(newItem, newMultiTaxonomyValues);
                        newItem.Update();
                    }

                }

                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Added {numOfItems} items to list {listTitle} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task<string> GetSiteFieldDefaultValue(ClientContext ctx, string fieldName)
        {
            try
            {
                Field field = ctx.Web.Fields.GetByTitle(fieldName);
                ctx.Load(field, f => f.DefaultValue);
                await ctx.ExecuteQueryAsync();
                return field.DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return string.Empty;
            }
        }

        public static async Task UpdateFieldDefaultValue(ClientContext ctx, string fieldName, string defaultValue, bool isTaxonomyField = false, string termSetName = "", string cityName = "")
        {
            // get field with fieldName
            try
            {
                if (isTaxonomyField)
                {
                    TaxonomyField field =
                        ctx.CastTo<TaxonomyField>(ctx.Site.RootWeb.Fields.GetByTitle(fieldName));
                    ctx.Load(field);
                    await ctx.ExecuteQueryAsync();

                    TermSet termSet = GetTermSet(ctx, termSetName);
                    Term cityTerm = termSet.Terms.GetByName(cityName);
                    ctx.Load(cityTerm);
                    await ctx.ExecuteQueryAsync();

                    TaxonomyFieldValue defaultTaxonomyValue = new TaxonomyFieldValue()
                    {
                        Label = cityTerm.Name,
                        TermGuid = cityTerm.Id.ToString(),
                        WssId = -1,
                    };
                    var validatedString = field.GetValidatedString(defaultTaxonomyValue);
                    field.DefaultValue = validatedString.Value;
                    field.UserCreated = false;
                    field.UpdateAndPushChanges(true);
                    await ctx.ExecuteQueryAsync();
                }
                else
                {
                    Field field = ctx.Site.RootWeb.Fields.GetByTitle(fieldName);
                    field.DefaultValue = defaultValue;
                    field.Update();
                    await ctx.ExecuteQueryAsync();
                }
                Console.WriteLine($"Update {fieldName} field site default value {defaultValue} successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task QueryNotAboutItemsCaml(ClientContext ctx, string listTitle)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View>
                                        <Query>
                                            <Where>
                                                <Neq>
                                                    <FieldRef Name='about' />
                                                    <Value Type='Text'>about default</Value>
                                                </Neq>
                                            </Where>
                                        </Query>
                                    </View>";
                ListItemCollection items = targetList.GetItems(camlQuery);
                ctx.Load(items);
                await ctx.ExecuteQueryAsync();
                foreach (ListItem item in items)
                {
                    Console.WriteLine(item["Title"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateListViewByCityOrderByCreatedTime(ClientContext ctx, string listTitle,
            string viewTitle, string cityName)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ViewCreationInformation viewCreationInfo = new ViewCreationInformation();
                viewCreationInfo.Title = viewTitle;
                viewCreationInfo.ViewTypeKind = ViewType.Html;
                viewCreationInfo.RowLimit = 10;
                viewCreationInfo.Paged = true;
                string commaSeperatedCols = "ID,Name,about,city";
                viewCreationInfo.ViewFields = commaSeperatedCols.Split(',');
                // query filter by city and order by created time
                viewCreationInfo.Query = $@"<Where>
                                                <Eq>
                                                    <FieldRef Name='city' />
                                                    <Value Type='TaxonomyFieldType'>{cityName}</Value>
                                                </Eq>
                                            </Where>
                                            <OrderBy>
                                                <FieldRef Name='Created' Ascending='False' />
                                            </OrderBy>";
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                View newView = targetList.Views.Add(viewCreationInfo);
                newView.Update();
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Created view {viewTitle} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        public static async Task UpdateBatchAboutColumnCaml(ClientContext ctx, string listTitle, string newAboutValue, int numOfItemsInBatch = 2)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                // create caml query to filter items with about value = about default
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View>
                                        <Query>
                                            <Where>
                                                <Eq>
                                                    <FieldRef Name='about' />
                                                    <Value Type='Text'>about default</Value>
                                                </Eq>
                                            </Where>
                                        </Query>
                                    </View>";

                ListItemCollection items = targetList.GetItems(camlQuery);
                ctx.Load(items);
                await ctx.ExecuteQueryAsync();

                int loopTimes = items.Count > numOfItemsInBatch ? numOfItemsInBatch : items.Count;

                for (int i = 0; i < loopTimes; i++)
                {
                    var item = items[i];
                    Console.WriteLine($"Value {item.Id} is put into update batch!");
                    item["about"] = newAboutValue;
                    item.UpdateOverwriteVersion();
                }

                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"{numOfItemsInBatch} updated!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // https://sharepoint.stackexchange.com/questions/275238/allowmultiplevalues-in-taxonomy-field-in-csom-javascript
        public static async Task CreateTaxonomyFieldMulti(ClientContext ctx, string fieldName,
            string termSetName)
        {
            try
            {
                var termSet = GetTermSet(ctx, termSetName);
                var taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
                var termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                ctx.Load(termStore);

                var createField = ctx.Web.Fields.AddFieldAsXml($@"
                        <Field Type='TaxonomyFieldTypeMulti' 
                             Name='{fieldName}' 
                             Mult='TRUE'
                             DisplayName='{fieldName}' StaticName='{fieldName}'
                             TermSetId='{{{termSet.Id.ToString()}}}'
                         />", true, AddFieldOptions.DefaultValue);

                ctx.Load(createField);
                await ctx.ExecuteQueryAsync();

                var updateTaxField = ctx.CastTo<TaxonomyField>(createField);
                updateTaxField.SspId = termStore.Id;
                updateTaxField.TermSetId = termSet.Id;
                updateTaxField.Update();
                await ctx.ExecuteQueryAsync();

                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Created field {fieldName} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateDocumentLibraryListType(ClientContext ctx, string listTitle)
        {
            try
            {
                ListCreationInformation listCreationInfo = new ListCreationInformation();
                listCreationInfo.Title = listTitle;
                listCreationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
                List newList = ctx.Web.Lists.Add(listCreationInfo);
                newList.Description = "This is a document library list";
                newList.ContentTypesEnabled = true;
                newList.EnableFolderCreation = true;
                newList.Update();
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Created document library list {listTitle} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // helper to create folder
        public static Folder EnsureFolder(ClientContext ctx, Folder parentFolder, string folderPath)
        {
            //Split up the incoming path so we have the first element as the a new sub-folder name 
            //and add it to parentFolder folders collection
            string[] pathElements = folderPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string head = pathElements[0];
            Folder newFolder = parentFolder.Folders.Add(head);
            ctx.Load(newFolder);
            ctx.ExecuteQuery();

            //If we have subfolders to create then the length of pathElements will be greater than 1
            if (pathElements.Length > 1)
            {
                //If we have more nested folders to create then reassemble the folder path using what we have left i.e. the tail
                string tail = string.Empty;
                for (int i = 1; i < pathElements.Length; i++)
                    tail = tail + "/" + pathElements[i];

                //Then make a recursive call to create the next subfolder
                return EnsureFolder(ctx, newFolder, tail);
            }
            else
                //This ensures that the folder at the end of the chain gets returned
                return newFolder;
        }

        public static async Task CreateFolderInList(ClientContext ctx, string listTitle, string folderName, string folderPathFromRoot = "")
        {
            try
            {

                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList.RootFolder);
                await ctx.ExecuteQueryAsync();
                EnsureFolder(ctx, targetList.RootFolder, folderPathFromRoot + "/" + folderName);
                Console.WriteLine($"Created folder {folderPathFromRoot + "/" + folderName} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateItemInFolder(ClientContext ctx, string listTitle, string folderUrl,
            string itemName, string about = "", string cities = "", string termSetName = "")
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                var serverRelativeUrl = ctx.Web.ServerRelativeUrl;
                ctx.Load(ctx.Web, w => w.Title);
                await ctx.ExecuteQueryAsync();

                var fileCreationInfo = new FileCreationInformation()
                {
                    Content = System.IO.File.ReadAllBytes(
                        "D:/projects/CSOM-Sharepoint-Exercise/ConsoleCSOM/Document1.docx"),
                    Url = $"{serverRelativeUrl}/{listTitle}{folderUrl}/{itemName}.docx"
                };
                var file = targetList.RootFolder.Files.Add(fileCreationInfo);
                ctx.Load(file);
                await ctx.ExecuteQueryAsync();

                ListItem newItem = file.ListItemAllFields;
                newItem["Title"] = itemName;
                newItem["about"] = about;
                newItem.Update();

                if (cities != "")
                {
                    // Todo: add multi tax value here: https://stackoverflow.com/questions/17076509/set-taxonomy-field-multiple-values
                    var multiTaxField = ctx.CastTo<TaxonomyField>(targetList.Fields.GetByInternalNameOrTitle("cities"));
                    var termCollection = GetTermSet(ctx, termSetName).Terms;
                    var listValue = cities.Split(',');
                    List<string> fieldValueList = new List<string>();
                    foreach (var item in listValue)
                    {
                        try
                        {
                            Term term = termCollection.GetByName(item);
                            ctx.Load(term);
                            await ctx.ExecuteQueryAsync();
                            string fieldValueItem = $"-1;#{term.Name}|{term.Id.ToString()}";
                            fieldValueList.Add(fieldValueItem);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            Console.WriteLine($"Error while loading term {newItem}!");
                        }
                    }
                    // README: https://sharepoint.stackexchange.com/questions/276504/the-given-value-for-a-taxonomy-field-was-not-formatted-in-the-required-intl
                    string fieldValue = string.Join(";#", fieldValueList);
                    TaxonomyFieldValueCollection newMultiTaxonomyValues =
                        new TaxonomyFieldValueCollection(ctx, fieldValue, multiTaxField);
                    multiTaxField.SetFieldValueByValueCollection(newItem, newMultiTaxonomyValues);
                    newItem.Update();
                }


                ctx.Load(newItem);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Created item {itemName} in folder {folderUrl} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // https://pholpar.wordpress.com/2018/03/23/how-to-check-if-a-specific-file-exists-in-a-folder-structure-of-a-sharepoint-document-library-using-the-client-object-model/
        public static async Task GetItemsWithStockHolmInFolderTask(ClientContext ctx, string folderUrl,
            string listTitle)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                var serverRelativeUrl = ctx.Web.ServerRelativeUrl;
                ctx.Load(ctx.Web, w => w.Title);
                await ctx.ExecuteQueryAsync();

                string folderServerRelativeUrl = string.Format("{0}/{1}{2}", serverRelativeUrl, listTitle, folderUrl);
                var camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"
                    <View>
                        <Query>
                            <Where>
                                <Contains>
                                    <FieldRef Name='cities' />
                                    <Value Type='TaxonomyFieldType'>Stockholm</Value>
                                </Contains>
                            </Where>
                        </Query>
                    </View>";
                camlQuery.FolderServerRelativeUrl = folderServerRelativeUrl;

                ListItemCollection items = targetList.GetItems(camlQuery);
                ctx.Load(items);
                await ctx.ExecuteQueryAsync();

                foreach (var item in items)
                {
                    Console.WriteLine(item["Title"]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateListFolderStructureOnlyTask(ClientContext ctx, string listTitle)
        {
            try
            {

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateFieldInListCsomHelperTask(ClientContext ctx, string listTitle,
            string fieldName, FieldType fieldType, string displayName)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                var createNewField = targetList.Fields.AddFieldAsXml($"<Field Type='{fieldType.ToString()}' DisplayName='{displayName}' Name='{fieldName}' StaticName='{fieldName}' />", true, AddFieldOptions.DefaultValue);
                ctx.Load(createNewField);
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Created field {fieldName} in list {listTitle} successfully!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        // best explanation for Ensure User: https://tutorials4sharepoint.wordpress.com/2018/07/11/ensureuser-method-in-sharepoint/
        public static async Task GetUserFromEmailOrName(ClientContext ctx, string emailOrName)
        {
            try
            {
                // get site users
                var web = ctx.Web;
                ctx.Load(web);
                await ctx.ExecuteQueryAsync();

                // get by name

                User user = web.EnsureUser(emailOrName);
                ctx.Load(user);
                await ctx.ExecuteQueryAsync();

                Console.WriteLine($"Get user {emailOrName} successfully!");
                // print out the user properties
                Console.WriteLine($"User: {user.Title} ----- Email: {user.Email} ---------- LoginName: {user.LoginName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static async Task CreateFolderViewAndMakeDefaultView(ClientContext ctx, string listTitle,
            string viewTitle)
        {
            try
            {
                List targetList = ctx.Web.Lists.GetByTitle(listTitle);
                ViewCreationInformation viewCreationInfo = new ViewCreationInformation();
                viewCreationInfo.Title = viewTitle;
                viewCreationInfo.ViewTypeKind = ViewType.Html;

                string commaSeperatedCols = "DocIcon,Name,FileDirRef";
                viewCreationInfo.ViewFields = commaSeperatedCols.Split(',');

                // https://sharepoint.stackexchange.com/questions/89844/caml-query-to-filter-items-by-content-type-independent-of-column-name
                viewCreationInfo.Query = @"<Where>
                                                <Eq>
                                                    <FieldRef Name='ContentType' />
                                                    <Value Type='Computed'>Folder</Value>
                                                </Eq>
                                            </Where>";
                ctx.Load(targetList);
                await ctx.ExecuteQueryAsync();

                View newView = targetList.Views.Add(viewCreationInfo);
                newView.DefaultView = true;
                newView.Update();
                await ctx.ExecuteQueryAsync();
                Console.WriteLine($"Created view {viewTitle} in list {listTitle} successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
