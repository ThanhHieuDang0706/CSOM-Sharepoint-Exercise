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

        public async Task UpdateUser(string accountName)
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
