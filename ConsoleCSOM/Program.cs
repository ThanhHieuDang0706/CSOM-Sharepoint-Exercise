using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using ConsoleCSOM.Search;
using ConsoleCSOM.UserProfile;
using Microsoft.SharePoint.Client.Taxonomy;


namespace ConsoleCSOM
{
    
    class Program
    {
       
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper())
                {
                    ClientContext ctx = Program.GetContext(clientContextHelper, "SharepointSearch");
                    SearchService searchService = new SearchService(ctx);

                    //await searchService.RunSearchQuery("contentclass:STS_ListItem_WebPageLibrary");
                    //await searchService.RunSearchQuery("cities:(StockHolm, Ho Chi Minh)");
                    //await searchService.RunSearchQuery("FirstNameOSWTEXT:Diego");

                    //await searchService.SearchInList("CSOM Test", "aboutOWSTEXT:\"Update Script\" city:\"Ho Chi Minh\"");

                    await searchService.SearchUser("AboutMe:\"Developer\"");

                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }

            Console.WriteLine("Press Any Key To Stop!");
            Console.ReadKey();
        }
        public static ClientContext GetContext(ClientContextHelper clientContextHelper, string key)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            var info = config.GetSection(key).Get<SharepointInfo>();
            Console.WriteLine($"{info.SiteUrl} -- {info.Username}");
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
    }
}
