using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
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
                    ClientContext ctx = Program.GetContext(clientContextHelper, "SharepointInfo");
                    UserProfileService userProfileService = new UserProfileService(ctx);

                    //await userProfileService.ListUserProperties();
                    await userProfileService.UpdatePersonTypeProperty("i:0#.f|membership|hieudang0706@zyntp.onmicrosoft.com", "SPS-DontSuggestList", "hieu.dang.thanh@preciofishbone.se");
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
            var info = config.GetSection("SharepointInfoUserPermissionExercise").Get<SharepointInfo>();
            Console.WriteLine($"{info.SiteUrl} -- {info.Username}");
            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
    }
}
