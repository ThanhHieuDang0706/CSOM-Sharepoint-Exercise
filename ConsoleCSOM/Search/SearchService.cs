using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;


namespace ConsoleCSOM.Search
{
    public class SearchService
    {
        private readonly ClientContext _ctx;

        public SearchService(ClientContext ctx)
        {
            _ctx = ctx;
        }

        public async Task RunSearchQuery(string query)
        {
            KeywordQuery keywordQuery = new KeywordQuery(_ctx);
            keywordQuery.QueryText = query;
            SearchExecutor searchExecutor = new SearchExecutor(_ctx);
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            await _ctx.ExecuteQueryAsync();
            string json = JsonSerializer.Serialize(results.Value, new JsonSerializerOptions() { WriteIndented = true });
            Console.WriteLine(json);
        }

        public async Task SearchInList(string listTitle, string query)
        {
            try
            {
                // get list
                List list = _ctx.Web.Lists.GetByTitle(listTitle);
                _ctx.Load(list, l => l.Id);
                await _ctx.ExecuteQueryAsync();

                KeywordQuery keywordQuery = new KeywordQuery(_ctx);
                keywordQuery.QueryText = $"(contentclass:STS_ListItem OR IsDocument:True) ListID:{list.Id} AND {query}";

                SearchExecutor searchExecutor = new SearchExecutor(_ctx);
                ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
                await _ctx.ExecuteQueryAsync();
                string json = JsonSerializer.Serialize(results.Value, new JsonSerializerOptions() { WriteIndented = true });
                Console.WriteLine(json);
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex.Message);
            }
        }

        public async Task SearchUser(string query)
        {
            try
            {
                KeywordQuery keywordQuery = new KeywordQuery(_ctx);
                keywordQuery.QueryText = query;
                keywordQuery.SourceId = Guid.Parse("b09a7990-05ea-4af9-81ef-edfab16c4e31");
                SearchExecutor searchExecutor = new SearchExecutor(_ctx);
                ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
                await _ctx.ExecuteQueryAsync();
                string json = JsonSerializer.Serialize(results.Value, new JsonSerializerOptions() { WriteIndented = true });

                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
