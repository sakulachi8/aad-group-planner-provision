using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http;

namespace Company.Function
{
    public static class planner
    {
        
        static readonly HttpClient client = new HttpClient();
        private static string activeDirectoryTenantId = Environment.GetEnvironmentVariable("activeDirectoryTenantId");
        private static string activeDirectoryClientId = Environment.GetEnvironmentVariable("activeDirectoryClientId");
        private static string activeDirectoryClientSecretId = Environment.GetEnvironmentVariable("activeDirectoryClientSecretId");
        [FunctionName("planner")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "planner")] HttpRequest req,
            ILogger log)
        {
            
            log.LogInformation("C# HTTP trigger function processed a request.");
            dynamic responseMessage = "Bad Request";
            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                dynamic data = JsonConvert.DeserializeObject(requestBody);
                string groupId = data?.GroupID;
                string plannerTitle = data?.Title;

                string activeDirectoryToken = await GetActiveDirectoryToken();
                Dictionary<string, dynamic> jsonData = new Dictionary<string, dynamic>()
                {
                    { "owner",groupId},
                    { "title",plannerTitle},
                };

                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + activeDirectoryToken);
                HttpResponseMessage httpResponseMessage = await client.PostAsJsonAsync("https://graph.microsoft.com/v1.0/planner/plans", jsonData);
                if (httpResponseMessage.IsSuccessStatusCode)
                {
                    Dictionary<string, dynamic> plannerData = await httpResponseMessage.Content.ReadAsAsync<Dictionary<string, dynamic>>();
                    return new OkObjectResult(plannerData);
                }
                    Dictionary<string, dynamic> groupData1 = await httpResponseMessage.Content.ReadAsAsync<Dictionary<string, dynamic>>();
            }
            catch(Exception ex)
            {
                responseMessage = responseMessage + ": " + ex.Message;
            }
            return new BadRequestObjectResult(responseMessage);




        }



        
        public static async Task<string> GetActiveDirectoryToken()
        {
            string result = null;
            Dictionary<string, string> jsonData = new Dictionary<string, string>()
            {
                { "grant_type","client_credentials"},
                { "client_id",activeDirectoryClientId},
                { "client_secret",activeDirectoryClientSecretId},
                { "resource","https://graph.microsoft.com"}
            };
            HttpResponseMessage responseActiveDirectory = await client.PostAsync("https://login.microsoftonline.com/" + activeDirectoryTenantId + "/oauth2/token", new FormUrlEncodedContent(jsonData));
            if (responseActiveDirectory.IsSuccessStatusCode)
            {
                result = (await responseActiveDirectory.Content.ReadAsAsync<Dictionary<string, dynamic>>())["access_token"];
            }
            return result;
        }

    }
}
