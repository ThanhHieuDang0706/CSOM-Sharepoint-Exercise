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
                    SearchService serachService = new SearchService(ctx);

                    //await serachService.RunSearchQuery("contentclass:STS_ListItem_WebPageLibrary");
                    //await serachService.RunSearchQuery("cities:(StockHolm, Ho Chi Minh)");
                    //await serachService.RunSearchQuery("Title:Developer");
                    await serachService.RunSearchQuery("FirstNameOWSTEXT:Diego");


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
