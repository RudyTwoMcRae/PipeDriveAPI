using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using System.Net;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using RestSharp.Deserializers;
using System.Collections;
using System.Xml;
using System.Threading;

namespace PipeDriveAPI
{
    public class Deals
    {
    #region Custom Classess used in various functions
        public class DynamicJsonDeserializer : IDeserializer
        {
            public string RootElement { get; set; }
            public string Namespace { get; set; }
            public string DateFormat { get; set; }

            public T Deserialize<T>(RestResponse response) where T : new()
            {
                return JsonConvert.DeserializeObject<dynamic>(response.Content);
            }

            public T Deserialize<T>(IRestResponse response)
            {
                throw new NotImplementedException();
            }
        }
        public class DealStatus
        {
            /// <summary>
            /// The Available Status's of a deal
            /// </summary>
            public enum dealStatus
            {
                Open = 1,
                Won = 2,
                Loss = 3
            }
            private dealStatus dsStatus = dealStatus.Open;

            /// <summary>
            /// Get or Set the status of the current deal
            /// </summary>
            public dealStatus Status
            {
                get { return dsStatus; }
                set { dsStatus = value; }
            }
            /// <summary>
            /// Return the string equavilent of Status
            /// </summary>
            public string StatusText
            {
                get
                {
                    switch (dsStatus)
                    {
                        case dealStatus.Open:
                            return "open";
                        case dealStatus.Won:
                            return "won";
                        case dealStatus.Loss:
                            return "loss";
                        default:
                            return "open";
                    }
                }
            }
        }
        public class UserAPI
        {
            string sUsername = "";

            public UserAPI(string name)
            {
                sUsername = name;
            }

            public string Username { get { return sUsername; } set { sUsername = value; } }
            public string PipeDriveUsername { get { return GetPipeDriveUserName(); } }
            public string APIKey { get { return GetAPIKey(); } }
            public int Stage { get { return GetStage(); } }

            public string GetAPIKey()
            {

                //Variables
                XmlDocument doc = new XmlDocument();
                doc.Load(@"\\192.168.1.7\McRae Files\PipeDrive\Users.xml");
                XmlNodeList xmlUsers = doc.GetElementsByTagName("user");
                //Deal Import Variables
                string sAPIKey = "";

                if (xmlUsers.Count <= 0)
                {
                    Console.WriteLine("ERROR:");
                    Console.WriteLine("There are no Users in the specified XML file.");
                    Console.ReadLine();
                }
                else
                {
                    //Find the User
                    foreach (XmlNode curUser in xmlUsers)
                    {
                        if (sUsername == curUser.SelectSingleNode("username").InnerText || sUsername == curUser.SelectSingleNode("shortname").InnerText)
                        {
                            sAPIKey = curUser.SelectSingleNode("key").InnerText;
                        }
                    }
                    
                }

                return sAPIKey;
            }
            public int GetStage()
            {

                //Variables
                XmlDocument doc = new XmlDocument();
                doc.Load(@"\\192.168.1.7\McRae Files\PipeDrive\Users.xml");
                XmlNodeList xmlUsers = doc.GetElementsByTagName("user");
                //Deal Import Variables
                int iStage = 9;

                if (xmlUsers.Count <= 0)
                {
                    Console.WriteLine("ERROR:");
                    Console.WriteLine("There are no Users in the specified XML file.");
                    Console.ReadLine();
                }
                else
                {
                    //Find the User
                    foreach (XmlNode curUser in xmlUsers)
                    {
                        if (sUsername == curUser.SelectSingleNode("username").InnerText || sUsername == curUser.SelectSingleNode("shortname").InnerText)
                        {
                            iStage = int.Parse(curUser.SelectSingleNode("stage").InnerText);
                        }
                    }

                }

                return iStage;
            }
            public string GetPipeDriveUserName()
            {

                //Variables
                XmlDocument doc = new XmlDocument();
                doc.Load(@"\\192.168.1.7\McRae Files\PipeDrive\Users.xml");
                XmlNodeList xmlUsers = doc.GetElementsByTagName("user");

                if (xmlUsers.Count <= 0)
                {
                    Console.WriteLine("ERROR:");
                    Console.WriteLine("There are no Users in the specified XML file.");
                    Console.ReadLine();
                    return null;
                }
                else
                {
                    //Find the User
                    foreach (XmlNode curUser in xmlUsers)
                    {
                        if (sUsername == curUser.SelectSingleNode("shortname").InnerText)
                        {
                            return curUser.SelectSingleNode("username").InnerText;
                        }
                    }
                    return null;
                }
            }
        }
    #endregion

    #region API Object Templates
        public class DealObject
        {
            //public int id { get; set; }
            public int org_id { get; set; }
            public int stage_id { get; set; }
            public int person_id { get; set; }
            public string value { get; set; }
            public string add_time { get; set; }
            public string title { get; set; }
            public string status { get; set; }
        }
        public class UpdateDealObject
        {
            public int id { get; set; }
            public int org_id { get; set; }
            public int stage_id { get; set; }
            public int person_id { get; set; }
            public string value { get; set; }
            public string add_time { get; set; }
            public string title { get; set; }
            public string status { get; set; }
        }
        public class OrganizationObject
        {
            public string name { get; set; }
            public int owner_id { get; set; }
            //public int visible_to { get; set; }
        }
        public class PersonsObject
        {
            public string name { get; set; }
            public int owner_id { get; set; }
            public int org_id { get; set; }
            public ArrayList email { get; set; }
            public ArrayList phone { get; set; }
            public int visible_to { get; set; }
        }
        public class EMail
        {
            public string label { get; set; }
            public string value { get; set; }
            public bool primary { get; set; }
        }
        public class Phone
        {
            public string label { get; set; }
            public string value { get; set; }
            public bool primary { get; set; }
        }
        #endregion

        #region Global Static Objects that are filled from FileMaker and used in API Calls
        private static DealObject g_Deal = new DealObject();
        private static UpdateDealObject g_UpdateDeal = new UpdateDealObject();
        private static OrganizationObject g_Org = new OrganizationObject();
        private static PersonsObject g_Person = new PersonsObject();
        private static UserAPI g_User = new UserAPI("Bob");
        #endregion

        #region Get API Calls - Mainly Used to get the ID's based on their name
        private static int GetOrgID(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/organizations/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.success == "true")
                {
                    dynamic dOrg = JObject.Parse(dResponse.data.First.ToString());
                    return dOrg.id;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("The organization with the name " + sName + "could not be found.");
                    Environment.Exit(-1);
                    return -1;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: GetOrgID()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return -1;
            }
        }
        private static bool OrgExists(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/organizations/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.data != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: OrgExists()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return false;
            }
        }
        private static int GetPersonID(string sName, string sOrgID)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/persons/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);
                rrRequest.AddQueryParameter("org_id", sOrgID);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.success == "true")
                {
                    dynamic dPerson = JObject.Parse(dResponse.data.First.ToString());
                    return dPerson.id;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("The person with the name " + sName + "could not be found.");
                    Environment.Exit(-1);
                    return -1;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: GetPersonID()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return -1;
            }
        }
        private static int GetPersonID(string sEmail)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/persons/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sEmail);
                rrRequest.AddQueryParameter("search_by_email", "1");

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.success == "true")
                {
                    dynamic dPerson = JObject.Parse(dResponse.data.First.ToString());
                    return dPerson.id;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("The person with the name " + sEmail + "could not be found.");
                    Environment.Exit(-1);
                    return -1;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: GetPersonID(string sEmail)");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return -1;
            }
        }
        private static bool PersonExists(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/persons/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.data != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: OrgExists()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return false;
            }
        }
        private static int GetDealID(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/deals/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.success == "true")
                {
                    dynamic dDeal = JObject.Parse(dResponse.data.First.ToString());
                    return dDeal.id;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("The deal with the name " + sName + "could not be found.");
                    Environment.Exit(-1);
                    return -1;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: GetDealID()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return -1;
            }
        }
        private static bool DealExists(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/deals/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.data != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: OrgExists()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return false;
            }
        }
        private static int GetOwnerID(string sName)
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/users/find?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.GET);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddQueryParameter("term", sName);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                if (dResponse.success == "true"  && dResponse.data != null)
                {
                    dynamic dUser = JObject.Parse(dResponse.data.First.ToString());
                    return dUser.id;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("The user with the name " + sName + "could not be found.");
                    Environment.Exit(-1);
                    return -1;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: GetOwnerID()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
                return -1;
            }
        }

        #endregion

        #region API Calls to Create/Update Things in PipeDrive
        public static void CreateDeal()
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            string sURL = @"https://mcrae2.pipedrive.com/v1/deals?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.POST);
            DealStatus dsStatus = new DealStatus { Status = DealStatus.dealStatus.Open };

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddJsonBody(g_Deal);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                string s = dResponse.data.ToString();
                dynamic dDeal = JObject.Parse(dResponse.data.ToString());

                Console.WriteLine();
                Console.WriteLine("Deal created successfully.");
                Console.WriteLine("Press Enter to close this window.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: CreateDeal()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }
        public static void ReloadDeal()
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            int iDealID = GetDealID(g_UpdateDeal.title);
            string sURL = @"https://mcrae2.pipedrive.com/v1/deals/" + iDealID + "?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.PUT);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddJsonBody(g_Deal);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                string s = dResponse.data.ToString();
                dynamic dDeal = JObject.Parse(dResponse.data.ToString());

                Console.WriteLine();
                Console.WriteLine("Deal Updated successfully.");
                Console.WriteLine("Press Enter to close this window.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: ReloadDeal()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                //Environment.Exit(-1);
            }
        }
        public static void UpdateDeal()
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            int iDealID = GetDealID(g_UpdateDeal.title);
            string sURL = @"https://mcrae2.pipedrive.com/v1/deals/" + iDealID + "?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.PUT);

            //Execute Request
            try
            {
                //Add Header and Parameters
                g_UpdateDeal.id = iDealID;
                rrRequest.AddJsonBody(g_UpdateDeal);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                string s = dResponse.data.ToString();
                dynamic dDeal = JObject.Parse(dResponse.data.ToString());

                Console.WriteLine();
                Console.WriteLine("Deal updated successfully.");
                Console.WriteLine("Press Enter to close this window.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: UpdateDeal()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                //Environment.Exit(-1);
            }
        }
        public static void CreateOrg()
        {
            //Variables
            string sAPIToken = g_User.APIKey;
            //sAPIToken = "24c8a9317e0f7714eac16a77089bb2a7eebb0204";
            string sURL = @"https://mcrae2.pipedrive.com/v1/organizations?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.POST);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddJsonBody(g_Org);
                //rrRequest.AddParameter("2eb2403b970222debf8d48478bf7efce95d62986",g_Org.name, ParameterType.RequestBody);
                //string sCustom = "{\"2eb2403b970222debf8d48478bf7efce95d62986\":\"" + g_Org.name + "\"";
                rrRequest.Parameters[0].Value = rrRequest.Parameters[0].Value.ToString().Substring(0, rrRequest.Parameters[0].Value.ToString().Length - 1) + ",\"2eb2403b970222debf8d48478bf7efce95d62986\":\"" + g_Org.name + "\"}";
                //rrRequest.AddParameter("application/json", sCustom, ParameterType.RequestBody);


                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                string s = dResponse.data.ToString();
                dynamic dOrg = JObject.Parse(dResponse.data.ToString());

                Console.WriteLine();
                Console.WriteLine("Organization created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: CreateOrg()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }
        public static void CreateCustomer()
        {
            //Variables
            string sAPIToken = "24c8a9317e0f7714eac16a77089bb2a7eebb0204";
            string sURL = @"https://mcrae2.pipedrive.com/v1/persons?api_token=" + sAPIToken;
            RestClient rcClient = new RestClient(sURL);
            RestRequest rrRequest = new RestRequest(Method.POST);

            //Execute Request
            try
            {
                //Add Header and Parameters
                rrRequest.AddJsonBody(g_Person);

                //Execute Request
                IRestResponse<dynamic> rrResponse = rcClient.Execute<dynamic>(rrRequest);

                //Get response from API
                dynamic dResponse = JObject.Parse(rrResponse.Content);
                string s = dResponse.data.ToString();
                dynamic dDeal = JObject.Parse(dResponse.data.ToString());

                Console.WriteLine();
                Console.WriteLine("Customer created successfully.");
                Console.WriteLine("Press Enter to close this window.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR:\n\tFunction: CreateCustomer()");
                Console.WriteLine("\t" + ex.Message);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }
        #endregion

        #region Load Functions
        public static bool LoadNewDeal(string sPath, string sStatus)
        {
            //Variables
            XmlDocument doc = new XmlDocument();
            doc.Load(sPath);
            XmlNodeList xmlJobs = doc.GetElementsByTagName("command");
            //Deal Import Variables
            string sOrgName = "";
            string sTitle = "";
            string sValue = "";
            string sRep = "";
            string sPerson = "";
            string sEmail = "";
            DealStatus dsStatus = new DealStatus { Status = DealStatus.dealStatus.Open };

            if (xmlJobs.Count <= 0)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine("There are no Jobs in the specified XML file.");
                Console.ReadLine();
                return true;
            }
            else
            {
                //Set the Deal Import Variables
                sOrgName = xmlJobs[0].SelectSingleNode("org_name").InnerText;
                sTitle = xmlJobs[0].SelectSingleNode("title").InnerText;
                sValue = xmlJobs[0].SelectSingleNode("value").InnerText;
                sRep = xmlJobs[0].SelectSingleNode("rep").InnerText;
                sPerson = xmlJobs[0].SelectSingleNode("client_name").InnerText;
                sEmail = xmlJobs[0].SelectSingleNode("email").InnerText;

                //Load global Deal object with values from FileMaker import
                g_User = new UserAPI(sRep);
                g_Deal.stage_id = g_User.GetStage();
                g_Deal.title = sTitle;
                g_Deal.value = sValue.Replace("$","");
                //g_Deal.org_id = GetOrgID(sOrgName);
                if (sEmail.Contains(";"))
                    sEmail = sEmail.Substring(0, sEmail.IndexOf(";"));
                g_Deal.person_id = GetPersonID(sEmail);
                //Hard Coded values for the Deal
                g_Deal.add_time = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                g_Deal.status = dsStatus.StatusText;

                if (DealExists(sTitle))
                    return false;
                else
                    return true;
            }
        }
        public static void LoadUpdateDeal(string sPath, string sStatus)
        {
            //Variables
            XmlDocument doc = new XmlDocument();
            doc.Load(sPath);
            XmlNodeList xmlJobs = doc.GetElementsByTagName("command");
            //Deal Import Variables
            string sOrgName = "";
            string sTitle = "";
            string sValue = "";
            string sRep = "";
            string sPerson = "";
            string sEmail = "";
            DealStatus dsStatus = new DealStatus { Status = DealStatus.dealStatus.Open };

            if (xmlJobs.Count <= 0)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine("There are no Jobs in the specified XML file.");
                Console.ReadLine();
            }
            else
            {
                //Set the Deal Import Variables
                sOrgName = xmlJobs[0].SelectSingleNode("org_name").InnerText;
                sTitle = xmlJobs[0].SelectSingleNode("title").InnerText;
                sValue = xmlJobs[0].SelectSingleNode("value").InnerText;
                sRep = xmlJobs[0].SelectSingleNode("rep").InnerText;
                sPerson = xmlJobs[0].SelectSingleNode("client_name").InnerText;
                sEmail = xmlJobs[0].SelectSingleNode("email").InnerText;

                //Load global Deal object with values from FileMaker import
                g_User = new UserAPI(sRep);
                g_UpdateDeal.stage_id = g_User.GetStage() + 1;
                g_UpdateDeal.title = sTitle;
                g_UpdateDeal.value = sValue.Replace("$", "");
                //g_UpdateDeal.org_id = GetOrgID(sOrgName);
                if (sEmail.Contains(";"))
                    sEmail = sEmail.Substring(0, sEmail.IndexOf(";"));
                g_Deal.person_id = GetPersonID(sEmail);
                g_UpdateDeal.person_id = GetPersonID(sEmail);
                //Hard Coded values for the Deal
                g_UpdateDeal.add_time = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                dsStatus.Status = DealStatus.dealStatus.Won;
                Console.WriteLine(dsStatus.StatusText);
                g_UpdateDeal.status = dsStatus.StatusText;
            }
        }
        public static void LoadNewOrg(string sPath)
        {
            //Variables
            XmlDocument doc = new XmlDocument();
            doc.Load(sPath);
            XmlNodeList xmlCustomer = doc.GetElementsByTagName("command");
            //Deal Import Variables
            string sOrg = "";
            string sUser = "";

            if (xmlCustomer.Count <= 0)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine("There are no Customers in the specified XML file.");
                Console.ReadLine();
            }
            else
            {
                //Set the Deal Import Variables
                Console.WriteLine(xmlCustomer[0].SelectSingleNode("org_name").InnerText);
                sOrg = xmlCustomer[0].SelectSingleNode("org_name").InnerText;
                sUser = xmlCustomer[0].SelectSingleNode("rep").InnerText;

                //Load the Global Org Object
                g_Org.name = sOrg;
                g_User = new UserAPI(sUser);
            }
        }
        public static void LoadNewCustomer(string sPath)
        {
            //Variables
            XmlDocument doc = new XmlDocument();
            doc.Load(sPath);
            XmlNodeList xmlCustomer = doc.GetElementsByTagName("command");
            //Deal Import Variables
            string sName = "";
            string sOrg = "";
            EMail eEMail = new EMail();
            Phone pPhone = new Phone();
            ArrayList arrEMails = new ArrayList();
            ArrayList arrPhone = new ArrayList();

            if (xmlCustomer.Count <= 0)
            {
                Console.WriteLine("ERROR:");
                Console.WriteLine("There are no Customers in the specified XML file.");
                Console.ReadLine();
            }
            else
            {
                //Set the Deal Import Variables
                sName = xmlCustomer[0].SelectSingleNode("client_name").InnerText;
                sOrg = xmlCustomer[0].SelectSingleNode("org_name").InnerText;
                eEMail.label = "work";
                eEMail.value = xmlCustomer[0].SelectSingleNode("email").InnerText;
                //Get only the first e-mail address
                if (eEMail.value.Contains(";"))
                    eEMail.value = eEMail.value.Substring(0, eEMail.value.IndexOf(";"));
                eEMail.primary = true;
                arrEMails.Add(eEMail);
                pPhone.label = "work";
                pPhone.value = xmlCustomer[0].SelectSingleNode("phone").InnerText;
                pPhone.primary = true;
                arrPhone.Add(pPhone);

                //Load global Deal object with values from FileMaker import
                g_Person.name = sName;
                //Determine if the orginization Exists and create it if it doen't
                LoadNewOrg(sPath);
                if (!OrgExists(g_Org.name))
                {
                    CreateOrg();
                    Thread.Sleep(5000);
                }
                g_Person.org_id = GetOrgID(sOrg);
                //g_Person.visible_to = iVisibleTo;
                g_Person.email = arrEMails;
                g_Person.phone = arrPhone;
                g_User = new UserAPI(xmlCustomer[0].SelectSingleNode("rep").InnerText);
                g_Person.owner_id = GetOwnerID(g_User.PipeDriveUsername);
                if (!PersonExists(eEMail.value))
                {
                    CreateCustomer();
                }

                Console.WriteLine(sName);
            }
        }
        #endregion
    }
}

