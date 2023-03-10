using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;

namespace ConsoleCSOM.UserProfile
{
    public class UserProfileService
    {
        private readonly PeopleManager _peopleManager;
        private readonly ClientContext _ctx;

        public UserProfileService(ClientContext ctx)
        {
            _peopleManager = new PeopleManager(ctx);
            _ctx = ctx;
        }

        public async Task ListUserProperties()
        {
            try
            {
                var userProperties = _peopleManager.GetMyProperties();
                _ctx.Load(userProperties);
                await _ctx.ExecuteQueryAsync();
                foreach (var key in userProperties.GetType().GetProperties())
                {
                    Console.WriteLine(key.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public async Task UpdateUserProperty(string accountName, string propertyName, string value)
        {
            try
            {
                var userProperties = _peopleManager.GetPropertiesFor(accountName);

                _ctx.Load(userProperties, u => u.AccountName);
                await _ctx.ExecuteQueryAsync();

                _peopleManager.SetSingleValueProfileProperty(userProperties.AccountName, propertyName, value);
                await _ctx.ExecuteQueryAsync();
                Console.WriteLine($"User {userProperties.AccountName} has been updated with property {propertyName} and value {value}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public async Task UpdatePersonTypeProperty(string accountName, string propertyName, string logonName)
        {
            try
            {
                UserCollection users = _ctx.Web.SiteUsers;
                _ctx.Load(users, u => u.Include(user => user.LoginName, user => user.Title, user => user.AadObjectId));

                User user = users.GetByEmail("hieudang0706@zyntp.onmicrosoft.com");
                _ctx.Load(user, u => u.LoginName, u => u.Title, u => u.AadObjectId);
                await _ctx.ExecuteQueryAsync();

                Console.WriteLine(user.LoginName);
                Console.WriteLine(user.AadObjectId.NameId);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public async Task UpdateUserMutipleChoicePropertyTask(string accountName)
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


    }
}
