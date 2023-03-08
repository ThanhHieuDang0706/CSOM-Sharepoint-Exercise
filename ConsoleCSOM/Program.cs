using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;


namespace ConsoleCSOM
{
    
    class Program
    {
       
        static async Task Main(string[] args)
        {
            try
            {
                //await CsomExerciseRunner.Run();
                await CsomUserPermissionExerciseRunner.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }

            Console.WriteLine("Press Any Key To Stop!");
            Console.ReadKey();
        }
    }
}
