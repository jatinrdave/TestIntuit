using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Forms;
using Intuit.Ipp.OAuth2PlatformClient;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using System.Security.Claims;
using System.IO;
using System.Reflection;
using Intuit.Ipp.DataService;
using System.Xml.Serialization;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    [Serializable]
    public class QboAuthTokens
    {
        [XmlElement("AccessToken")]
        public string AccessToken { get; set; }
        [XmlElement("AccessTokenExpireAt")]
        public string AccessTokenExpireAt { get; set; }
        [XmlElement("RefreshToken")]
        public string RefreshToken { get; set; }
        [XmlElement("RefreshTokenExpireAt")]
        public string RefreshTokenExpireAt { get; set; }
        [XmlElement("RealmId")]
        public string RealmId { get; set; }
        [XmlElement("ClientId")]
        public string ClientId { get; set; } = "#############################################";
        [XmlElement("ClientSecret")]
        public string ClientSecret { get; set; } = "########################################";
        [XmlElement("RedirectUrl")]
        public string RedirectUrl { get; set; } = "http://localhost:27353/callback";
        [XmlElement("Environment")]
        public string Environment { get; set; } = "sandbox";
    }
    public partial class Form1 : Form
    {
        public QboAuthTokens Tokens = new QboAuthTokens();
        public OAuth2Client auth2Client = null;
        public Form1()
        {
            InitializeComponent();
        }
        public void saveTokenInfo()
        {
            var serializer = new XmlSerializer(typeof(QboAuthTokens));
            using (var writer = new System.IO.StreamWriter("Tokens.xml"))
            {
                serializer.Serialize(writer, Tokens);
            }
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
            // Initialize the WebView2 control
            // and begin the authorization
            // process inside the WebView2 control.
            if (System.IO.File.Exists("Tokens.xml"))
            {
                var serializer = new XmlSerializer(typeof(QboAuthTokens));
                using (var reader = new System.IO.StreamReader("Tokens.xml"))
                {
                    Tokens = (QboAuthTokens)serializer.Deserialize(reader);
                }
            }
            auth2Client = new OAuth2Client(Tokens.ClientId, Tokens.ClientSecret, Tokens.RedirectUrl, Tokens.Environment);
            if (Tokens.AccessToken == null || Tokens.AccessToken == "" )
            {
                this.RunAuthorization();
            }
            else
            {
                if (DateTime.Parse(Tokens.RefreshTokenExpireAt) <= DateTime.Now)
                    this.RunAuthorization();
                else
                {
                    if(await RefreshAccessToken())
                        this.label1.Text = "Authentication not needed close the form";
                    else
                        this.RunAuthorization();
                }
            }
        }
        public async Task<bool> RefreshAccessToken()
        {
            try
            {
                TokenResponse tokenResponse = await auth2Client.RefreshTokenAsync(Tokens.RefreshToken);
                Tokens.AccessToken = tokenResponse.AccessToken;
                Tokens.AccessTokenExpireAt = (DateTime.Now.AddSeconds(tokenResponse.AccessTokenExpiresIn)).ToString();
                saveTokenInfo();
            }
            catch(Exception e)
            {
                return false;
            }
            return true;
        }

        public async void RunAuthorization()
        {
            List < OidcScopes > scopes = new List<OidcScopes>();
            scopes.Add(OidcScopes.Accounting);
            string authorizeUrl = auth2Client.GetAuthorizationURL(scopes);
            //var r = Process.Start(authorizeUrl);
            // Initialize the WebView2 control.
            await webView21.EnsureCoreWebView2Async(null);

            // Navigate the WebView2 control to
            // a generated authorization URL.
            webView21.CoreWebView2.Navigate(authorizeUrl);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // When the user closes the form we
            // assume that the operation has
            // completed with success or failure.

            // Get the current query parameters
            // from the current WebView source (page).

            // Use the the shared helper library
            // to validate the query parameters
            // and write the output file.
            bool setSucessfully = true;
            if (Tokens.AccessToken == null || Tokens.AccessToken == "")
            {
                setSucessfully = false;
                string query = webView21.Source.Query;
                setSucessfully = CheckQueryParamsAndSet(query);
            }
            if (setSucessfully && Tokens.AccessToken != null && Tokens.AccessToken != "")
            {
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(Tokens.AccessToken);

                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(Tokens.RealmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "23";
                    serviceContext.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/"; //This is sandbox Url. Change to Prod Url if you are using production

                    // Create a QuickBooks QueryService using ServiceContext
                    QueryService<CompanyInfo> querySvc = new QueryService<CompanyInfo>(serviceContext);
                    QueryService<Customer> querySvc2 = new QueryService<Customer>(serviceContext);
                   
                    CompanyInfo companyInfo = querySvc.ExecuteIdsQuery("SELECT * FROM CompanyInfo").FirstOrDefault();
                    var allCustomer = querySvc2.ExecuteIdsQuery("SELECT * FROM Customer where Id='5'");          
                    string output = "Company Name: " + companyInfo.CompanyName + " Company Address: " + companyInfo.CompanyAddr.Line1 + ", " + companyInfo.CompanyAddr.City + ", " + companyInfo.CompanyAddr.Country + " " + companyInfo.CompanyAddr.PostalCode;
                    Console.WriteLine(output);
                    Console.WriteLine("Total customer with id = 5 : "+Convert.ToString( allCustomer.Count));
                /*
                if (allCustomer.Count > 0)
                {
                    Customer first = allCustomer.First();
                    Type type = first.GetType();
                    FieldInfo[] fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance);
                    PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

                    Console.WriteLine("Fields and Values:");

                    foreach (FieldInfo field in fields)
                    {
                        object value = field.GetValue(first);
                        Console.WriteLine($"{field.Name}: {value}");
                    }

                    Console.WriteLine("\nProperties and Values:");

                    foreach (PropertyInfo property in properties)
                    {
                        object value = property.GetValue(first);
                        Console.WriteLine($"{property.Name}: {value}");
                    }
                }
                */
                //addCustomer(serviceContext);
                //editCustomer(serviceContext);
                //QueryItems(serviceContext);
                QueryAccount(serviceContext);
                //AddItem(serviceContext);

            }
            else
            {
                MessageBox.Show("Quickbooks Online failed to authenticate.");
            }
        }

        /// <summary>
        /// Checks the passed <paramref name="queryString"/>.
        /// <br/>
        /// If the query was successful, the function returns <c>true</c> and sets the Token values.
        /// <br/>
        /// Otherwise the function returns <c>false</c> or throws an exception when <paramref name="suppressErrors"/> is <c>false</c>.
        /// </summary>
        /// <param name="queryString"></param>
        /// <param name="suppressErrors"></param>
        /// <returns></returns>
        /// <exception cref="InvalidDataException"></exception>
        public bool CheckQueryParamsAndSet(string queryString, bool suppressErrors = true)
        {
            // Parse the query string into a
            // NameValueCollection for easy access
            // to each parameter.
            Dictionary<string, string> query = ParseQueryString(queryString);

            // Make sure the required query
            // parameters exist.
            if (query["code"] != null && query["realmId"] != null)
            {

                // Use the OAuth2Client to get a new
                // access token from the QBO servers.
                TokenResponse responce = auth2Client.GetBearerTokenAsync(query["code"]).Result;

                // Set the token values with the client
                // responce and query parameters.
                Tokens.AccessToken = responce.AccessToken;
                Tokens.RefreshToken = responce.RefreshToken;
                Tokens.RealmId = query["realmId"];
                Tokens.AccessTokenExpireAt = (DateTime.Now.AddSeconds(responce.AccessTokenExpiresIn)).ToString();
                Tokens.RefreshTokenExpireAt = (DateTime.Now.AddSeconds(responce.RefreshTokenExpiresIn)).ToString();
                saveTokenInfo();
                // Return true. The Tokens have
                // been set as expected.
                return true;
            }
            else
            {

                // Is the caller chooses to suppress
                // errors return false instead
                // of throwing an exception.
                if (suppressErrors)
                {
                    return false;
                }
                else
                {
                    throw new InvalidDataException(
                        $"The 'code' or 'realmId' was not present in the query parameters '{query}'."
                    );
                }
            }
        }
        static Dictionary<string, string> ParseQueryString(string url)
        {
            Dictionary<string, string> queryParameters = new Dictionary<string, string>();
            int questionMarkIndex = url.IndexOf('?');

            if (questionMarkIndex != -1)
            {
                string queryString = url.Substring(questionMarkIndex + 1);
                string[] parts = queryString.Split('&');

                foreach (string part in parts)
                {
                    string[] keyValue = part.Split('=');
                    if (keyValue.Length == 2)
                    {
                        string key = Uri.UnescapeDataString(keyValue[0]);
                        string value = Uri.UnescapeDataString(keyValue[1]);
                        queryParameters[key] = value;
                    }
                }
            }

            return queryParameters;
        }
        /*private async System.Threading.Tasks.Task GetAuthTokensAsync(string code, string realmId)
        {
            var tokenResponse = await auth2Client.GetBearerTokenAsync(code);

            if (realmId != null && realmId!="")
            {
                claims["realmId"] = realmId;
            }
            if (!string.IsNullOrWhiteSpace(tokenResponse.AccessToken))
            {
                claims["access_token"] = tokenResponse.AccessToken;
                claims["access_token_expires_at"]= (DateTime.Now.AddSeconds(tokenResponse.AccessTokenExpiresIn)).ToString();
            }

            if (!string.IsNullOrWhiteSpace(tokenResponse.RefreshToken))
            {
                claims["refresh_token"] = tokenResponse.RefreshToken;
                claims["refresh_token_expires_at"] = (DateTime.Now.AddSeconds(tokenResponse.RefreshTokenExpiresIn)).ToString();
            }
        }*/

        /*private async void button2_Click(object sender, EventArgs e)
        {
            // Parse the URL to extract the query parameters
            Dictionary<string, string> queryParams = ParseQueryString(this.textBox1.Text.ToString());

            // Extract the values of code, state, realmId, and error (if present)
            queryParams.TryGetValue("code", out string code);
            queryParams.TryGetValue("state", out string state);
            queryParams.TryGetValue("realmId", out string realmId);
            queryParams.TryGetValue("error", out string error);
            await GetAuthTokensAsync(code, realmId);
            try
            {
                if (claims.TryGetValue("access_token", out string access_token))
                {
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(access_token);

                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "69";

                    // Create a QuickBooks QueryService using ServiceContext
                    QueryService<CompanyInfo> querySvc = new QueryService<CompanyInfo>(serviceContext);
                    var ans = querySvc.ExecuteIdsQuery("SELECT * FROM CompanyInfo");
                    CompanyInfo companyInfo = null;
                    //if (ans!=null)
                        //companyInfo= ()

                    //this.label1.Text = "Company Name: " + companyInfo.CompanyName + " Company Address: " + companyInfo.CompanyAddr.Line1 + ", " + companyInfo.CompanyAddr.City + ", " + companyInfo.CompanyAddr.Country + " " + companyInfo.CompanyAddr.PostalCode;
                }
                else
                    this.label1.Text = "Invalid access tokens";
            }
            catch (Exception ex)
            {
                this.label1.Text = "QBO API call Failed!" + " Error message: " + ex.Message;
            }
        }*/

        public void addCustomer(ServiceContext serviceContext)
        {

            Customer newCustomer = new Customer
            {
                DisplayName = "Sample Customer 1",
                GivenName = "Sample",
                FamilyName = "Customer 1",
                PrimaryEmailAddr = new EmailAddress { Address = "samplecustoomer1@email.com" }
            };
            DataService dataService = new DataService(serviceContext);

            // Add the customer to QuickBooks Online
            Customer addedCustomer = dataService.Add<Customer>(newCustomer);
            if (addedCustomer != null)
            {
                Console.WriteLine("Customer added successfully. Customer ID: " + addedCustomer.Id);
            }
            else
            {
                Console.WriteLine("Failed to add the customer.");
            }
        }

        public void editCustomer(ServiceContext serviceContext)
        {
            QueryService<Customer> queryService = new QueryService<Customer>(serviceContext);
            Customer existingCustomer = queryService.ExecuteIdsQuery($"SELECT * FROM Customer WHERE Id = '59'").FirstOrDefault();
            existingCustomer.PrimaryEmailAddr = new EmailAddress { Address = "samplecustomer1@email.com" };
            DataService dataService = new DataService(serviceContext);
            // Update the customer in QuickBooks Online
            Customer updatedCustomer = dataService.Update(existingCustomer);
            if (updatedCustomer != null)
            {
                Console.WriteLine("Customer updated successfully. Customer ID: " + updatedCustomer.Id);
            }
            else
            {
                Console.WriteLine("Failed to update the customer.");
            }
        }

        public void QueryItems(ServiceContext serviceContext)
        {
            QueryService<Item> querySvc = new QueryService<Item>(serviceContext);
            var allItems = querySvc.ExecuteIdsQuery("SELECT * FROM Item");
            foreach(Item i1 in allItems)
            {
                Console.WriteLine(i1.Id + i1.Name+ i1.ItemCategoryType);
            }
        }
        public void AddItem(ServiceContext serviceContext)
        {
            Item newItem = new Item();
            newItem.Name = "SY01";
            newItem.Type = ItemTypeEnum.NonInventory;
            newItem.IncomeAccountRef=null;
            DataService dataService = new DataService(serviceContext);
            Item addedItem = dataService.Add<Item>(newItem);
            if (addedItem != null)
                Console.WriteLine("Item addeed sucess fully Id: " + addedItem.Id + " Type: " + addedItem.Type + " Name:" + addedItem.Name + " QtyOnHand: " + addedItem.QtyOnHand);
            else
                Console.WriteLine("Failed to add Item .....!!");

        }

        public void QueryAccount(ServiceContext serviceContext)
        {
            QueryService<Account> qsrv = new QueryService<Account>(serviceContext);
            var allAccount = qsrv.ExecuteIdsQuery("Select * from Account");
            foreach (Account i1 in allAccount)
            {
                Console.WriteLine(i1.Id + i1.Name );
            }
        }
        public void CreateInvoice(ServiceContext serviceContext)
        {
            Invoice newInvoice = new Invoice();
            newInvoice.CurrencyRef = new ReferenceType { Value = "59"};
            List<Line> listOfLines = new List<Line>();
        }
    }
}
