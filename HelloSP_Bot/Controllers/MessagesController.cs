using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Utilities;
using Newtonsoft.Json;
using Microsoft.SharePoint;
using System.Xml;

namespace HelloSP_Bot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        SPWeb Web;
        string StatusValue;
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<Message> Post([FromBody]Message message)
        {


            if (message.Type == "Message")
            {
                // calculate something for us to return
                int length = (message.Text ?? string.Empty).Length;


                // return our reply to the user
                //return message.CreateReplyMessage("Hello World !!!");
                string NewsString;
                NewsLUIS StLUIS = await GetEntityFromLUIS(message.Text);
                if (StLUIS.intents.Count() > 0)
                {
                    switch (StLUIS.intents[0].intent)
                    {
                        case "GetTopNews":
                            NewsString = GetListData(StLUIS.entities[0].entity);
                            break;
                        case "GoToGoogle":
                            NewsString = "Please click on hyperlink to search more https://www.google.com/#q=" + message.Text;
                            //await GetListData(StLUIS.entities[0].entity);
                            break;
                        default:
                            NewsString = "Sorry, I am still process of learning human. Try in other words..";
                            break;
                    }
                }
                else
                {
                    NewsString = "Sorry, I am not getting you...";
                }

                // return our reply to the user  
                return message.CreateReplyMessage(NewsString);

            }
            else
            {
                return HandleSystemMessage(message);
            }
        }

        private static async Task<NewsLUIS> GetEntityFromLUIS(string Query)
        {
            Query = Uri.EscapeDataString(Query);
            NewsLUIS Data = new NewsLUIS();
            using (HttpClient client = new HttpClient())
            {
                string RequestURI = "https://api.projectoxford.ai/luis/v1/application?id=908f2c82-436a-4ee1-b931-076db449168e&subscription-key=4360ab34eaae45c8b9506a4c75e9e1e7&q=" + Query;
                HttpResponseMessage msg = await client.GetAsync(RequestURI);

                if (msg.IsSuccessStatusCode)
                {
                    var JsonDataResponse = await msg.Content.ReadAsStringAsync();
                    Data = JsonConvert.DeserializeObject<NewsLUIS>(JsonDataResponse);
                }
            }
            return Data;
        }



        public SPWeb GetWeb()
        {
            try
            {

                using (SPSite site = new SPSite("http://spwfedevv002:6907/sites/Importad/"))
                {
                    using (Web = site.OpenWeb())
                    {

                    }
                }
                return Web;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public string GetListData(object count)
        {




            /*Declare and initialize a variable for the Lists Web service.*/
            Web_Reference.Lists listService = new Web_Reference.Lists();

            /*Authenticate the current user by passing their default 
            credentials to the Web service from the system credential cache.*/
            listService.Credentials =
                System.Net.CredentialCache.DefaultCredentials;

            /*Set the Url property of the service for the path to a subsite.*/
            listService.Url =
                "http://spwfedevv002:6907/sites/Importad/_vti_bin/lists.asmx";

            // Instantiate an XmlDocument object
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();

            /* Assign values to the string parameters of the GetListItems method, using GUIDs for the listName and viewName variables. For listName, using the list display name will also work, but using the list GUID is recommended. For viewName, only the view GUID can be used. Using an empty string for viewName causes the default view 
            to be used.*/
            string listName = "{6C3EEF09-72D7-4BBD-857E-3E368DCB5859}";
            string viewName = "{062B24E9-C4BD-4B2D-8550-3529394970D2}";
            string rowLimit = "150";

            /*Use the CreateElement method of the document object to create elements for the parameters that use XML.*/
            System.Xml.XmlElement query = xmlDoc.CreateElement("Query");
            System.Xml.XmlElement viewFields =
                xmlDoc.CreateElement("ViewFields");
            System.Xml.XmlElement queryOptions =
                xmlDoc.CreateElement("QueryOptions");

            /*To specify values for the parameter elements (optional), assign CAML fragments to the InnerXml property of each element.*/
            query.InnerXml = "<Where></Where>";
            viewFields.InnerXml = "<FieldRef Name=\"Title\" />";
            queryOptions.InnerXml = "";

            /* Declare an XmlNode object and initialize it with the XML response from the GetListItems method. The last parameter specifies the GUID of the Web site containing the list. Setting it to null causes the Web site specified by the Url property to be used.*/
            System.Xml.XmlNode Result =
                listService.GetListItems
                (listName, viewName, query, viewFields, rowLimit, queryOptions, null);

            /*Loop through each node in the XML response and display each item.*/


            XmlDocument NewxmlDoc = new XmlDocument();

            string ListItemsNamespacePrefix = "z";
            string ListItemsNamespaceURI = "#RowsetSchema";

            string PictureLibrariesNamespacePrefix = "s";
            string PictureLibrariesNamespaceURI = "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882";

            string WebPartsNamespacePrefix = "dt";
            string WebPartsNamespaceURI = "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882";

            string DirectoryNamespacePrefix = "rs";
            string DirectoryNamespaceURI = "urn:schemas-microsoft-com:rowset";

            //now associate with the xmlns namespaces (part of all XML nodes returned 
            //from SharePoint) a namespace prefix which we can then use in the queries 
            System.Xml.XmlNamespaceManager nsmgr = new System.Xml.XmlNamespaceManager(NewxmlDoc.NameTable);

            nsmgr.AddNamespace(ListItemsNamespacePrefix, ListItemsNamespaceURI);
            nsmgr.AddNamespace(PictureLibrariesNamespacePrefix, PictureLibrariesNamespaceURI);
            nsmgr.AddNamespace(WebPartsNamespacePrefix, WebPartsNamespaceURI);
            nsmgr.AddNamespace(DirectoryNamespacePrefix, DirectoryNamespaceURI);


            XmlNodeList nodeList = Result.SelectNodes("//z:row", nsmgr);


            int test = 1;
            foreach (XmlNode node in nodeList)
            {

                StatusValue += test.ToString() + '-' + node.Attributes["ows_Title"].InnerText;
                test++;

            }

            return StatusValue;

        }
        public void GetOneListItemValue()
        {

            try
            {
                SPList EmployeeList = Web.Lists.TryGetList("Sections");
                if (EmployeeList != null)
                {
                    SPListItem EmployeeListItem = EmployeeList.Items.GetItemById(1);


                    if (EmployeeListItem["Title"] != null)
                    {
                        StatusValue = EmployeeListItem["Title"].ToString();


                    }

                }

            }
            catch (Exception)
            {

                throw;
            }

        }

        private Message HandleSystemMessage(Message message)
        {
            if (message.Type == "Ping")
            {
                Message reply = message.CreateReplyMessage();
                reply.Type = "Ping";
                return reply;
            }
            else if (message.Type == "DeleteUserData")
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == "BotAddedToConversation")
            {
            }
            else if (message.Type == "BotRemovedFromConversation")
            {
            }
            else if (message.Type == "UserAddedToConversation")
            {
            }
            else if (message.Type == "UserRemovedFromConversation")
            {
            }
            else if (message.Type == "EndOfConversation")
            {
            }

            return null;
        }
    }
}