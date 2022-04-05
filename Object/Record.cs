using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using ZCRMSDK.CRM.Library.Api.Response;
using ZCRMSDK.CRM.Library.CRMException;
using ZCRMSDK.CRM.Library.CRUD;
using ZCRMSDK.CRM.Library.Setup.RestClient;
using ZCRMSDK.OAuth.Client;
using ZCRMSDK.OAuth.Contract;

namespace WebservicesSage.Object
{
        class Record
        {
            public Record()
            {
                Dictionary<string, string> zcrmConfigurations = new Dictionary<string, string>(){
                {"client_id","1000.xxxxxxx"},{"client_secret","xxxxxxxxxx"},
                {"redirect_uri","https://www.zoho.com"},{"persistence_handler_class","ZCRMSDK.OAuth.ClientApp.ZohoOAuthFilePersistence, ZCRMSDK"},
                {"oauth_tokens_file_path","token.txt"},{"logFilePath","LogFile.txt"}
            };
                ZCRMRestClient.Initialize(zcrmConfigurations); ZCRMRestClient.SetCurrentUser("usermail@domain.com");
                string userEmailID = zcrmConfigurations.ContainsKey("currentUserEmail") ? zcrmConfigurations["currentUserEmail"] : ZCRMRestClient.GetCurrentUserEmail();
                if (!IsTokenGenerated(userEmailID))
                {
                    ZohoOAuthTokens tokens = ZohoOAuthClient.GetInstance().GenerateAccessToken("paste_the_self_authorized_grant_token_here");
                    Console.WriteLine("access token = " + tokens.AccessToken + " & refresh token = " + tokens.RefreshToken);
                }
            }
            public static bool IsTokenGenerated(string currentUserEmail)
            {
                try
                {
                    ZohoOAuthTokens tokens = ZohoOAuth.GetPersistenceHandlerInstance().GetOAuthTokens(currentUserEmail);
                    return tokens != null ? true : false;
                }
                catch (Exception)
                {
                    return false;
                }
            }
            /** Insert a specific record */
            public void InsertRecord()
            {
                try
                {
                    ZCRMRecord recordIns = new ZCRMRecord("Invoices"); //To get ZCRMRecord instance
                    recordIns.SetFieldValue("Subject", "Invoice"); //This method use to set FieldApiName and value similar to all other FieldApis and Custom field 
                    recordIns.SetFieldValue("Account_Name", "test");
                    recordIns.SetFieldValue("Customfield", "CustomFieldValue");
                    /** Following methods are being used only by Inventory modules */
                    ZCRMTax linetax = ZCRMTax.GetInstance("Sales Tax");
                    linetax.Percentage = 12.5;
                    recordIns.AddTax(linetax);
                    ZCRMRecord product1 = ZCRMRecord.GetInstance("Products", 34770601); // product instance
                    ZCRMInventoryLineItem lineItem1 = new ZCRMInventoryLineItem(product1)
                    {
                        Description = "Product_description", //To set line item description
                        Discount = 5, //To set line item discount
                        ListPrice = 100 //To set line item list price
                    }; //To get ZCRMInventoryLineItem instance
                    lineItem1.DiscountPercentage = 10;
                    ZCRMTax taxInstance11 = ZCRMTax.GetInstance("Sales Tax"); //To get ZCRMTax instance
                    taxInstance11.Percentage = 2; //To set tax percentage
                    taxInstance11.Value = 50; //To set tax value
                    lineItem1.AddLineTax(taxInstance11); //To set line tax to line item
                    lineItem1.Quantity = 500; //To set product quantity to this line item
                    recordIns.AddLineItem(lineItem1); //The line item set to the record object
                                                      /** End Inventory **/
                    List<string> trigger = new List<string>() { "workflow", "approval", "blueprint" };
                    APIResponse response = recordIns.Create(trigger); //To call the Update record method
                    ZCRMRecord record = (ZCRMRecord)response.Data;
                    Console.WriteLine("EntityId:" + record.EntityId); //To get update record id
                    Console.WriteLine("HTTP Status Code:" + response.HttpStatusCode); //To get Update record http response code
                    Console.WriteLine("Status:" + response.Status); //To get Update record response status
                    Console.WriteLine("Message:" + response.Message); //To get Update record response message
                    Console.WriteLine("Details:" + response.ResponseJSON); //To get Update record response details
                }
                catch (ZCRMException ex)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(ex));
                }
            }
            /*static void Main(string[] args)
            {
                Record record = new Record();
                record.InsertRecord();
            }*/
        }
    }
