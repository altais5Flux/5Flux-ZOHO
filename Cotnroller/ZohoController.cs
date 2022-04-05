using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZCRMSDK.OAuth.Contract;
using ZCRMSDK;
using ZCRMSDK.CRM.Library.Setup.RestClient;
using ZCRMSDK.OAuth.Client;
using ZCRMSDK.CRM.Library.CRUD;
using ZCRMSDK.CRM.Library.Api.Response;
using System.Net;
using System.IO;
using System.Web.Script.Serialization;
using ZCRMSDK.CRM.Library.Setup.Users;
using WebservicesSage.Object;
using Newtonsoft.Json;
using ZCRMSDK.CRM.Library.CRMException;
using Objets100cLib;
using WebservicesSage.Singleton;
using System.Configuration;
using System.Windows.Forms;
using ZCRMSDK.CRM.Library.Common;
using WebservicesSage.Utils;

namespace WebservicesSage.Cotnroller
{
    public class ZohoController
    {
        public static Dictionary<string, string> config = new Dictionary<string, string>()
        {
            {"client_id", ConfigurationManager.AppSettings["client_id"]},
            {"client_secret", ConfigurationManager.AppSettings["client_secret"]},
            {"redirect_uri", "http://www.google.com"},
            {"iamUrl","https://accounts.zoho.eu" },
            {"access_type", "offline"},
            {"loginAuthClass", "ZCRMSDK.CRM.Library.Common.ZCRMConfigUtil, ZCRMSDK"},
            {"persistence_handler_class", "ZCRMSDK.OAuth.ClientApp.ZohoOAuthFilePersistence, ZCRMSDK"},
            {"oauth_tokens_file_path", @ConfigurationManager.AppSettings["oauth_tokens_file_path"]},
            {"apiBaseUrl", "https://www.zohoapis.eu"},
            {"photoUrl", "https://profile.zoho.eu/api/v1/user/self/photo"},
            {"apiVersion", "v2"},
            {"logFilePath", ""},
            {"timeout", ""},
            {"minLogLevel", "ALL"},
            {"domainSuffix", "eu"},
            {"currentUserEmail", ConfigurationManager.AppSettings["currentUserEmail"]}

        };
        
        public ZohoController(string test)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + " :  restart App" + Environment.NewLine);
                File.AppendAllText("Log\\reboot.txt", sb.ToString());
                sb.Clear();
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ZCRMRestClient.Initialize(config);
            ZohoOAuthClient client = ZohoOAuthClient.GetInstance();
                string grantToken = ConfigurationManager.AppSettings["grantToken"];// "1000.48f3cbb7cec8fdbb65a64508fbbc03cf.b09910683129df7765fb4fea88151827";
                
                if (ConfigurationManager.AppSettings["useGrant"].Equals("TRUE"))
                {
                    try
                    {
                    ZohoOAuthTokens tokens = client.GenerateAccessToken(grantToken);
                    var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    var settings = configFile.AppSettings.Settings;
                    settings["useGrant"].Value = "FALSE";
                    configFile.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                    }
                    catch (ZCRMException ex)
            {
                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + ex.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Code " + ex.Code + Environment.NewLine);
                s.Append("IsAPIException " + ex.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + ex.IsSDKException + Environment.NewLine);
                s.Append("Message " + ex.Message + Environment.NewLine);
                s.Append("ErrorDetails " + ex.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + ex.ErrorMsg + Environment.NewLine);
                File.AppendAllText("Log\\grant.txt", s.ToString());

            }
                    
        }
            }
            catch (ZCRMException ex)
            {
                Console.WriteLine(ex.HttpStatusCode);
                Console.WriteLine(ex.Code);
                Console.WriteLine(ex.IsAPIException);
                Console.WriteLine(ex.IsSDKException);
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.ErrorDetails);
                Console.WriteLine(ex.ErrorMsg);
            }
        }
        public static void UpdateClient(Client account)
        {
            try
            {

            
            ZCRMModule module1 = ZCRMModule.GetInstance("Accounts");
                /*BulkAPIResponse<ZCRMRecord> response1 = module1.GetRecord(Int64.Parse(account.ZohoEntityId));// SearchByCriteria("CT_Num:equals:" + account.CT_NUM + "");
            List<ZCRMRecord> Clientrecords = response1.BulkData;*/

                APIResponse response1 = module1.GetRecord(Int64.Parse(account.ZohoEntityId));// SearchByCriteria("CT_Num:equals:" + account.CT_NUM + "");
                ZCRMRecord Client = (ZCRMRecord)response1.Data;
               /* foreach (ZCRMRecord Client in Clientrecords)
            {*/
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                IBOClient3 clientsSageObj = compta.FactoryClient.ReadNumero(account.CT_NUM);
                DateTime test = DateTime.Parse(Client.ModifiedTime);
                DateTime test2 = clientsSageObj.DateModification;
                // test pour savoire la plus récente date de modification dans sage ou bien dans le crm
                if (DateTime.Parse(Client.ModifiedTime) > clientsSageObj.DateModification)
                {
                    //Mise a jour données dans Sage
                    if (!Client.Data["Account_Name"].ToString().Equals(clientsSageObj.CT_Intitule))
                    {
                        clientsSageObj.CT_Intitule = Client.Data["Account_Name"].ToString();
                    }
                    /*if (!Client.Data["Phone"].ToString().Equals(clientsSageObj.Telecom.Telephone))
                    {
                        clientsSageObj.Telecom.Telephone = Client.Data["Phone"].ToString();
                    }*/
                    /*if (!Client.Data["Fax"].ToString().Equals(clientsSageObj.Telecom.Telecopie))
                    {
                        clientsSageObj.Telecom.Telecopie = Client.Data["Phone"].ToString();
                    }*/
                    if (!Client.Data["Code_NAF_Sage"].ToString().Equals(clientsSageObj.CT_Ape))
                    {
                        clientsSageObj.CT_Ape = Client.Data["Code_NAF_Sage"].ToString();
                    }
                    if (!Client.Data["Compte_Collectif"].ToString().Equals(clientsSageObj.CompteGPrinc.CG_Intitule))
                    {
                        clientsSageObj.CompteGPrinc.CG_Intitule = Client.Data["Compte_Collectif"].ToString();
                    }
                    if (!Client.Data["E_mail"].ToString().Equals(clientsSageObj.Telecom.EMail))
                    {
                        clientsSageObj.Telecom.EMail = Client.Data["E_mail"].ToString();
                    }
                    if (!Client.Data["Facebook"].ToString().Equals(clientsSageObj.CT_Facebook))
                    {
                        clientsSageObj.CT_Facebook = Client.Data["Facebook"].ToString();
                    }
                    if (!Client.Data["LinkedIn"].ToString().Equals(clientsSageObj.CT_LinkedIn))
                    {
                        clientsSageObj.CT_LinkedIn = Client.Data["LinkedIn"].ToString();
                    }
                    if (!Client.Data["Interlocuteur"].ToString().Equals(clientsSageObj.CT_Contact))
                    {
                        clientsSageObj.CT_Contact = Client.Data["Interlocuteur"].ToString();
                    }
                    if (!Client.Data["Identifiant_TVA"].ToString().Equals(clientsSageObj.CT_Identifiant))
                    {
                        clientsSageObj.CT_Identifiant = Client.Data["Identifiant_TVA"].ToString();
                    }
                    if (!Client.Data["Qualite"].ToString().Equals(clientsSageObj.CT_Qualite))
                    {
                        clientsSageObj.CT_Qualite = Client.Data["Qualite"].ToString();
                    }
                    if (!Client.Data["Website"].ToString().Equals(clientsSageObj.Telecom.Site))
                    {
                        clientsSageObj.Telecom.Site = Client.Data["Website"].ToString();
                    }
                    if (!Client.Data["Commentaire"].ToString().Equals(clientsSageObj.CT_Commentaire))
                    {
                        clientsSageObj.CT_Commentaire = Client.Data["Commentaire"].ToString();
                    }
                    if (!Client.Data["Groupe_Tarifaire"].ToString().Equals(clientsSageObj.CatTarif.CT_Intitule))
                    {
                        clientsSageObj.CatTarif.CT_Intitule = Client.Data["Groupe_Tarifaire"].ToString();
                    }
                    if (!Client.Data["SIRET"].ToString().Equals(clientsSageObj.CT_Siret))
                    {
                        clientsSageObj.CT_Siret = Client.Data["SIRET"].ToString();
                    }
                    clientsSageObj.Write();
                }
                else
                {
                    Client client = new Client(clientsSageObj);
                    ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                    ZCRMRecord IRecord = new ZCRMRecord("Accounts");
                    ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//USER WHO CREATED THE API TOKEN
                    IRecord.Owner = user;
                    IRecord.SetFieldValue("id", Int64.Parse(account.ZohoEntityId));
                    IRecord.SetFieldValue("Account_Name", account.Intitule);
                    IRecord.SetFieldValue("Phone", account.telephone);
                    if (account.clientBillingAdresses != null)
                    {
                        IRecord.SetFieldValue("Billing_Street", account.clientBillingAdresses[0].Adresse);
                        IRecord.SetFieldValue("Billing_City", account.clientBillingAdresses[0].Ville);
                        IRecord.SetFieldValue("Billing_State", account.clientBillingAdresses[0].Region);
                        IRecord.SetFieldValue("Billing_Code", account.clientBillingAdresses[0].CodePostal);
                        IRecord.SetFieldValue("Billing_Country", account.clientBillingAdresses[0].Pays);
                    }
                    IRecord.SetFieldValue("CT_Num", account.CT_NUM);
                    IRecord.SetFieldValue("Fax", account.telecopie);
                    IRecord.SetFieldValue("Code_NAF_Sage", account.CodeNAF);
                    IRecord.SetFieldValue("Compte_Collectif	", account.CompteCollectif);
                    IRecord.SetFieldValue("E_mail", account.Email);
                    IRecord.SetFieldValue("Facebook", account.Facebook);
                    IRecord.SetFieldValue("LinkedIn", account.LinkedIn);
                    IRecord.SetFieldValue("Interlocuteur", account.Interlocuteur);
                    IRecord.SetFieldValue("Identifiant_TVA", account.IdTVA);
                    IRecord.SetFieldValue("Qualite", account.qualite);
                    IRecord.SetFieldValue("Website", account.Website);
                    IRecord.SetFieldValue("Commentaire", account.commentaire);
                    IRecord.SetFieldValue("Groupe_Tarifaire", account.GroupeTarifaireIntitule);
                    IRecord.SetFieldValue("SIRET", account.Siret);
                    IRecord.SetFieldValue("Centrale_d_achat", account.CentralAchat);
                    IRecord.SetFieldValue("Repr_sentant", account.representant);
                    IRecord.SetFieldValue("Sommeil", account.Sommeil);
                    IRecord.SetFieldValue("Groupe", account.groupe);
                    IRecord.SetFieldValue("Enseigne", account.enseigne);

                    List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
                    ZCRMModule module = restClient1.GetModuleInstance("Accounts");
                    BulkAPIResponse<ZCRMRecord> response = module.UpdateRecords(records);
                    foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
                    {
                        Console.WriteLine((eResponse.ResponseJSON["details"]["id"].ToString()));//Fetches Created Account Id
                        Console.WriteLine(eResponse.ResponseJSON);//Fetches response as JSON
                        Console.WriteLine(eResponse.Status);//Fetches status value present in the response
                        Console.WriteLine(eResponse.Code);//Fetches code value present in the response
                        Console.WriteLine(eResponse.Message);//Fetches message value present in the response

                    }
                }
            //}

            }
            
            catch (ZCRMException Exception)
            {

                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Code " + Exception.Code + Environment.NewLine);
                s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                s.Append("Message " + Exception.Message + Environment.NewLine);
                s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                File.AppendAllText("Log\\syncClient.txt", s.ToString());
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + "ct_num : " + account.CT_NUM + Environment.NewLine);
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\syncClient.txt", sb.ToString());
                sb.Clear();
            }

        }
        public static void CreateProspect(ZCRMRecord Client)
        {
            try
            {
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                IBOClient3 clientsSageObj = (IBOClient3)SingletonConnection.Instance.Compta.FactoryProspect.Create();
                if (String.IsNullOrEmpty(Client.Data["CT_Num"].ToString()))
                {
                    string name = Client.Data["First_Name"].ToString() + Client.Data["Last_Name"].ToString();
                    clientsSageObj.CT_Intitule = name;

                    int iterID = 0;
                    while (compta.FactoryProspect.ExistNumero(UtilsConfig.PrefixClient + iterID.ToString()))
                    {
                        iterID++;
                    }
                    clientsSageObj.CT_Num = UtilsConfig.PrefixClient + iterID.ToString();
                    clientsSageObj.Write();
                    if (name.Length > 17)
                    {
                        clientsSageObj.CT_Classement = name.Substring(0, 17);
                    }
                    else
                    {
                        clientsSageObj.CT_Classement = name;
                    }
                    clientsSageObj.Telecom.Telephone = Client.Data["Phone"].ToString();
                    clientsSageObj.Telecom.Telecopie = Client.Data["Fax"].ToString();
                    clientsSageObj.CT_Ape = Client.Data["Code_NAF_Sage"].ToString();
                    clientsSageObj.Telecom.EMail = Client.Data["Email"].ToString();
                    clientsSageObj.CT_Facebook = Client.Data["Facebook"].ToString();
                    clientsSageObj.CT_LinkedIn = Client.Data["LinkedIn"].ToString();
                    clientsSageObj.CT_Contact = Client.Data["Interlocuteur"].ToString();
                    clientsSageObj.CT_Identifiant = Client.Data["Identifiant_TVA"].ToString();
                    clientsSageObj.CT_Qualite = Client.Data["Qualite"].ToString();
                    //clientsSageObj.Telecom.Site = Client.Data["Website"].ToString();
                    clientsSageObj.CT_Commentaire = Client.Data["Commentaire"].ToString();
                    clientsSageObj.CT_Siret = Client.Data["SIRET"].ToString();
                    clientsSageObj.Adresse.Adresse = Client.Data["Street"].ToString();
                    clientsSageObj.Adresse.CodePostal = Client.Data["Zip_Code"].ToString();
                    clientsSageObj.Adresse.Ville = Client.Data["City"].ToString();
                    clientsSageObj.Adresse.Pays = Client.Data["Country"].ToString();

                    clientsSageObj.Write();
                    DB.Update(Client.EntityId.ToString(), clientsSageObj.CT_Num);
                }

                ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                Client.Data["CT_Num"] = clientsSageObj.CT_Num;
                ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//USER WHO CREATED THE API TOKEN
               // Client.Owner = user;
                Client.SetFieldValue("id", Int64.Parse(Client.EntityId.ToString()));
                List<ZCRMRecord> records = new List<ZCRMRecord> { Client };
                ZCRMModule module = restClient1.GetModuleInstance("Leads");
                BulkAPIResponse<ZCRMRecord> response = module.UpdateRecords(records);
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + "ZOHO CRM LEAD ID : "+ Client.Data["id"].ToString() + Environment.NewLine);
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\SageCreateProspect.txt", sb.ToString());
                sb.Clear();
            }

        }
        public static void UpdateProspect(Prospect account)
        {
            try
            {
                ZCRMModule module1 = ZCRMModule.GetInstance("Leads");
                /*
                BulkAPIResponse<ZCRMRecord> response1 = module1.SearchByCriteria("CT_Num:equals:" + account.CT_NUM);
                List<ZCRMRecord> Clientrecords = response1.BulkData;
                foreach (ZCRMRecord Client in Clientrecords)
                {*/
                APIResponse response1 = module1.GetRecord(Int64.Parse(account.ZohoEntityId));// SearchByCriteria("CT_Num:equals:" + account.CT_NUM + "");
                ZCRMRecord Client = (ZCRMRecord)response1.Data;
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                    IBOClient3 clientsSageObj = compta.FactoryProspect.ReadNumero(account.CT_NUM);
                    DateTime test = DateTime.Parse(Client.ModifiedTime);
                    DateTime test2 = clientsSageObj.DateModification;
                    //test pour savoire la plus récente date de modification dans sage ou bien dans le crm
                    if (DateTime.Parse(Client.ModifiedTime) > clientsSageObj.DateModification)
                    {
                        //Mise a jour données dans Sage
                        if (!Client.Data["Company"].ToString().Equals(clientsSageObj.CT_Intitule))
                        {
                            clientsSageObj.CT_Intitule = Client.Data["Company"].ToString();
                        }
                        if (!Client.Data["Phone"].ToString().Equals(clientsSageObj.Telecom.Telephone))
                        {
                            clientsSageObj.Telecom.Telephone = Client.Data["Phone"].ToString();
                        }
                        if (!Client.Data["Fax"].ToString().Equals(clientsSageObj.Telecom.Telecopie))
                        {
                            clientsSageObj.Telecom.Telecopie = Client.Data["Fax"].ToString();
                        }
                        if (!Client.Data["Code_NAF_Sage"].ToString().Equals(clientsSageObj.CT_Ape))
                        {
                            clientsSageObj.CT_Ape = Client.Data["Code_NAF_Sage"].ToString();
                        }
                        /*if (!Client.Data["Compte_Collectif"].ToString().Equals(clientsSageObj.CompteGPrinc.CG_Intitule))
                        {
                            clientsSageObj.CompteGPrinc.CG_Intitule = Client.Data["Compte_Collectif"].ToString();
                        }*/
                        if (!Client.Data["E_mail"].ToString().Equals(clientsSageObj.Telecom.EMail))
                        {
                            clientsSageObj.Telecom.EMail = Client.Data["E_mail"].ToString();
                        }
                        if (!Client.Data["Facebook"].ToString().Equals(clientsSageObj.CT_Facebook))
                        {
                            clientsSageObj.CT_Facebook = Client.Data["Facebook"].ToString();
                        }
                        if (!Client.Data["LinkedIn"].ToString().Equals(clientsSageObj.CT_LinkedIn))
                        {
                            clientsSageObj.CT_LinkedIn = Client.Data["LinkedIn"].ToString();
                        }
                        if (!Client.Data["Interlocuteur"].ToString().Equals(clientsSageObj.CT_Contact))
                        {
                            clientsSageObj.CT_Contact = Client.Data["Interlocuteur"].ToString();
                        }
                        if (!Client.Data["Identifiant_TVA"].ToString().Equals(clientsSageObj.CT_Identifiant))
                        {
                            clientsSageObj.CT_Identifiant = Client.Data["Identifiant_TVA"].ToString();
                        }
                        if (!Client.Data["Qualite"].ToString().Equals(clientsSageObj.CT_Qualite))
                        {
                            clientsSageObj.CT_Qualite = Client.Data["Qualite"].ToString();
                        }
                        /*if (!Client.Data["Website"].ToString().Equals(clientsSageObj.Telecom.Site))
                        {
                            clientsSageObj.Telecom.Site = Client.Data["Website"].ToString();
                        }*/
                        if (!Client.Data["Commentaire"].ToString().Equals(clientsSageObj.CT_Commentaire))
                        {
                            clientsSageObj.CT_Commentaire = Client.Data["Commentaire"].ToString();
                        }
                        /*if (!Client.Data["Groupe_Tarifaire"].ToString().Equals(clientsSageObj.CatTarif.CT_Intitule))
                        {
                            clientsSageObj.CatTarif.CT_Intitule = Client.Data["Groupe_Tarifaire"].ToString();
                        }*/
                        if (!Client.Data["SIRET"].ToString().Equals(clientsSageObj.CT_Siret))
                        {
                            clientsSageObj.CT_Siret = Client.Data["SIRET"].ToString();
                        }
                        clientsSageObj.Write();
                    }
                    else
                    {
                        Prospect client = new Prospect(clientsSageObj);
                        ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                        ZCRMRecord IRecord = new ZCRMRecord("Leads");
                        ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//20069463844);//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
                        IRecord.Owner = user;
                        IRecord.SetFieldValue("id", Int64.Parse(account.ZohoEntityId));
                        IRecord.SetFieldValue("Company", account.Intitule);
                        IRecord.SetFieldValue("Phone", account.telephone);
                        if (account.clientBillingAdresses.Count > 0)
                        {
                            IRecord.SetFieldValue("Street", account.clientBillingAdresses[0].Adresse);
                            IRecord.SetFieldValue("City", account.clientBillingAdresses[0].Ville);
                            IRecord.SetFieldValue("Zip_Code", account.clientBillingAdresses[0].CodePostal);
                            IRecord.SetFieldValue("Country", account.clientBillingAdresses[0].Pays);
                        }
                        IRecord.SetFieldValue("CT_Num", account.CT_NUM);
                        IRecord.SetFieldValue("Fax", account.telecopie);
                        IRecord.SetFieldValue("Code_NAF_Sage", account.CodeNAF);
                        //IRecord.SetFieldValue("Compte_Collectif	", account.CompteCollectif);
                        IRecord.SetFieldValue("Email", account.Email);
                        IRecord.SetFieldValue("Facebook", account.Facebook);
                        IRecord.SetFieldValue("LinkedIn", account.LinkedIn);
                        IRecord.SetFieldValue("Interlocuteur", account.Interlocuteur);
                        IRecord.SetFieldValue("Identifiant_TVA", account.IdTVA);
                        IRecord.SetFieldValue("Qualite", account.qualite);
                        //IRecord.SetFieldValue("Website", account.Website);
                        IRecord.SetFieldValue("Commentaire", account.commentaire);
                        IRecord.SetFieldValue("Groupe_Tarifaire", account.GroupeTarifaireIntitule);
                        IRecord.SetFieldValue("SIRET", account.Siret);
                        IRecord.SetFieldValue("Centrale_d_achat", account.CentralAchat);
                        IRecord.SetFieldValue("Representant", account.representant);
                        IRecord.SetFieldValue("Sommeil", account.Sommeil);
                        if (account.Contacts.Count > 0)
                        {
                            IRecord.SetFieldValue("Nom_Contact", account.Contacts[0].Nom);
                            IRecord.SetFieldValue("Prenom_Contact", account.Contacts[0].Prenom);
                            IRecord.SetFieldValue("Telecopie_Contact", account.Contacts[0].Telecopie);
                            IRecord.SetFieldValue("Telephone_Contact", account.Contacts[0].Telephone);
                            IRecord.SetFieldValue("Portable_Contact", account.Contacts[0].Portable);
                            IRecord.SetFieldValue("Civilite_Contact", account.Contacts[0].Civilite);
                            IRecord.SetFieldValue("Skype_Contact", account.Contacts[0].Skype);
                            IRecord.SetFieldValue("Facebook_Contact", account.Contacts[0].Facebook);
                            IRecord.SetFieldValue("Fonction_Contact", account.Contacts[0].Fonction);
                            IRecord.SetFieldValue("Service_Contact", account.Contacts[0].Service);
                            IRecord.SetFieldValue("Email_Contact", account.Contacts[0].Email);
                            IRecord.SetFieldValue("LinkedIn_Contact", account.Contacts[0].LinkedIn);
                        }

                        List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
                        ZCRMModule module = restClient1.GetModuleInstance("Leads");
                        BulkAPIResponse<ZCRMRecord> response = module.UpdateRecords(records);

                    }
                //}
            }
            
            catch (ZCRMException ex)
            {
                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + ex.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Code " + ex.Code + Environment.NewLine);
                s.Append("IsAPIException " + ex.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + ex.IsSDKException + Environment.NewLine);
                s.Append("Message " + ex.Message + Environment.NewLine);
                s.Append("ErrorDetails " + ex.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + ex.ErrorMsg + Environment.NewLine);
                File.AppendAllText("Log\\updateprospectError.txt", s.ToString());
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\updateprospectError.txt", sb.ToString());
                sb.Clear();
            }
        }
        public static void SendClient(Client account)
        {
            try

            {


                ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                ZCRMRecord IRecord = new ZCRMRecord("Accounts");
                ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//20069463844);//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
                IRecord.Owner = user;
                IRecord.SetFieldValue("Account_Name", account.Intitule);
                IRecord.SetFieldValue("Phone", account.telephone);
                if (account.clientBillingAdresses != null)
                {
                    IRecord.SetFieldValue("Billing_Street", account.clientBillingAdresses[0].Adresse);
                    IRecord.SetFieldValue("Billing_City", account.clientBillingAdresses[0].Ville);
                    IRecord.SetFieldValue("Billing_State", account.clientBillingAdresses[0].Region);
                    IRecord.SetFieldValue("Billing_Code", account.clientBillingAdresses[0].CodePostal);
                    IRecord.SetFieldValue("Billing_Country", account.clientBillingAdresses[0].Pays);
                }
                IRecord.SetFieldValue("CT_Num", account.CT_NUM);
                IRecord.SetFieldValue("Fax", account.telecopie);
                IRecord.SetFieldValue("Code_NAF_Sage", account.CodeNAF);
                IRecord.SetFieldValue("Compte_Collectif	", account.CompteCollectif);
                IRecord.SetFieldValue("E_mail", account.Email);
                IRecord.SetFieldValue("Facebook", account.Facebook);
                IRecord.SetFieldValue("LinkedIn", account.LinkedIn);
                IRecord.SetFieldValue("Interlocuteur", account.Interlocuteur);
                IRecord.SetFieldValue("Identifiant_TVA", account.IdTVA);
                IRecord.SetFieldValue("Qualite", account.qualite);
                IRecord.SetFieldValue("Website", account.Website);
                IRecord.SetFieldValue("Commentaire", account.commentaire);
                IRecord.SetFieldValue("Groupe_Tarifaire", account.GroupeTarifaireIntitule);
                IRecord.SetFieldValue("SIRET", account.Siret);
                IRecord.SetFieldValue("Centrale_d_achat", account.CentralAchat);
                IRecord.SetFieldValue("Repr_sentant", account.representant);
                IRecord.SetFieldValue("Sommeil", account.Sommeil);
                IRecord.SetFieldValue("Groupe", account.groupe);
                IRecord.SetFieldValue("Enseigne", account.enseigne);
                List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
                ZCRMModule module = restClient1.GetModuleInstance("Accounts");
                BulkAPIResponse<ZCRMRecord> response = module.CreateRecords(records);
                String idProspec = "";
                foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
                {
                    idProspec = eResponse.ResponseJSON["details"]["id"].ToString();
                }
                try
                {
                    IBOClient3 clientSage = SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(account.CT_NUM);
                    var infolibreField = Singleton.SingletonConnection.Instance.Compta.FactoryTiers.InfoLibreFields;
                    int compteur = 1;

                    foreach (var infoLibreValue in clientSage.InfoLibre)
                    {
                        if (infolibreField[compteur].Name.Equals("ZohoEntityID"))
                        {
                            clientSage.InfoLibre[compteur] = idProspec;
                            break;
                        }
                        compteur++;
                    }
                    //clientSage.InfoLibre['6'] = eResponse.ResponseJSON["details"]["id"].ToString();
                    clientSage.Write();
                }
                catch (Exception s)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\zohoEntityId.txt", sb.ToString());
                    sb.Clear();
                }
                foreach (Contact item in account.Contacts)
                {
                    ZCRMRecord Record = new ZCRMRecord("Contacts");
                    Record.Owner = user;
                    Record.SetFieldValue("Last_Name", item.Nom);
                    Record.SetFieldValue("First_Name", item.Prenom);
                    Record.SetFieldValue("Fax", item.Telecopie);
                    Record.SetFieldValue("Phone", item.Telephone);
                    Record.SetFieldValue("Mobile", item.Portable);
                    //Record.SetFieldValue("Salutation", item.Civilite); not exist
                    //Record.SetFieldValue("Skype_Contact", item.Skype);
                    Record.SetFieldValue("Facebook", item.Facebook);
                    Record.SetFieldValue("Fonction", item.Fonction);
                    Record.SetFieldValue("Service", item.Service);
                    Record.SetFieldValue("Email", item.Email);
                    Record.SetFieldValue("Account_Name", account.Intitule);

                    List<ZCRMRecord> records1 = new List<ZCRMRecord>();// { Record };
                    records1.Add(Record);
                        //APIResponse resINsert = Record.Create();
                        ZCRMModule module1 = restClient1.GetModuleInstance("Contacts");
                        BulkAPIResponse<ZCRMRecord> response1 = module1.CreateRecords(records1); //module1.CreateRecords(records1);
                        foreach (EntityResponse eRespons in response1.BulkEntitiesResponse)
                        {
                            Console.WriteLine(eRespons.ResponseJSON["id"]);//Fetches Created Account Id
                            Console.WriteLine(eRespons.ResponseJSON);//Fetches response as JSON
                            Console.WriteLine(eRespons.Status);//Fetches status value present in the response
                            Console.WriteLine(eRespons.Code);//Fetches code value present in the response
                            Console.WriteLine(eRespons.Message);//Fetches message value present in the response
                        }

                }
            }
            catch (ZCRMException ex)
            {
                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + ex.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Code " +ex.Code + Environment.NewLine);
                s.Append("IsAPIException " +ex.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " +ex.IsSDKException + Environment.NewLine);
                s.Append("Message " +ex.Message + Environment.NewLine);
                s.Append("ErrorDetails " +ex.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " +ex.ErrorMsg + Environment.NewLine);
                File.AppendAllText("Log\\SendClient.txt", s.ToString());
                
            }
        }
        public static void SendProspect(Prospect account)
        {

            string id = "";
            ZCRMRestClient restClient = ZCRMRestClient.GetInstance();
            ZCRMRecord IRecord = new ZCRMRecord("Leads");
            ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
            IRecord.Owner = user;
            IRecord.SetFieldValue("Company", account.Intitule);
            IRecord.SetFieldValue("Last_Name", account.Intitule);
            IRecord.SetFieldValue("Phone", account.telephone);
            IRecord.SetFieldValue("CT_Num", account.CT_NUM);
            IRecord.SetFieldValue("Facebook", account.Facebook);
            IRecord.SetFieldValue("LinkedIn", account.LinkedIn);
            IRecord.SetFieldValue("Interlocuteur", account.Interlocuteur);
            IRecord.SetFieldValue("Identifiant_TVA", account.IdTVA);
            //IRecord.SetFieldValue("Website", account.Website);
            IRecord.SetFieldValue("Sommeil", account.Sommeil);
            IRecord.SetFieldValue("Email", account.Email);
            IRecord.SetFieldValue("Commentaire", account.commentaire);
            IRecord.SetFieldValue("Groupe_Tarifaire", account.GroupeTarifaireIntitule);
            IRecord.SetFieldValue("Fax", account.telecopie);
            if (account.Contacts.Count > 0)
            {
                IRecord.SetFieldValue("Nom_Contact", account.Contacts[0].Nom);
                IRecord.SetFieldValue("Prenom_Contact", account.Contacts[0].Prenom);
                IRecord.SetFieldValue("Telecopie_Contact", account.Contacts[0].Telecopie);
                IRecord.SetFieldValue("Telephone_Contact", account.Contacts[0].Telephone);
                IRecord.SetFieldValue("Portable_Contact", account.Contacts[0].Portable);
                IRecord.SetFieldValue("Civilite_Contact", account.Contacts[0].Civilite);
                IRecord.SetFieldValue("Skype_Contact", account.Contacts[0].Skype);
                IRecord.SetFieldValue("Facebook_Contact", account.Contacts[0].Facebook);
                IRecord.SetFieldValue("Fonction_Contact", account.Contacts[0].Fonction);
                IRecord.SetFieldValue("Service_Contact", account.Contacts[0].Service);
                IRecord.SetFieldValue("Email_Contact", account.Contacts[0].Email);
                IRecord.SetFieldValue("LinkedIn_Contact", account.Contacts[0].LinkedIn);
                IRecord.SetFieldValue("E_mail", account.Contacts[0].Email);

            }
            if (account.clientBillingAdresses.Count > 0)
            {
                IRecord.SetFieldValue("Street", account.clientBillingAdresses[0].Adresse);
                IRecord.SetFieldValue("City", account.clientBillingAdresses[0].Ville);
                //IRecord.SetFieldValue("Billing_State", account.clientBillingAdresses[0].Region);
                IRecord.SetFieldValue("Zip_Code", account.clientBillingAdresses[0].CodePostal);
                IRecord.SetFieldValue("Country", account.clientBillingAdresses[0].Pays);
            }
            IRecord.SetFieldValue("SIRET", account.Siret);
            IRecord.SetFieldValue("Centrale_d_achat", account.CentralAchat);
            IRecord.SetFieldValue("Representant", account.representant);
            IRecord.SetFieldValue("Code_NAF_Sage", account.CodeNAF);
            IRecord.SetFieldValue("Qualite", account.qualite);
            IRecord.SetFieldValue("Groupe", account.groupe);
            IRecord.SetFieldValue("Enseigne", account.enseigne);

            try
            {
                List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
                ZCRMModule module = restClient.GetModuleInstance("Leads");
                BulkAPIResponse<ZCRMRecord> response = module.CreateRecords(records);
                string idProspec = "";
                foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
                {
                    idProspec = eResponse.ResponseJSON["details"]["id"].ToString();
                }
                DB.Update(idProspec, account.CT_NUM);

            }
            catch (ZCRMException Exception)
            {
                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Source " + Exception.Source + Environment.NewLine);
                s.Append("HelpLink " + Exception.HelpLink + Environment.NewLine);
                s.Append("Code " + Exception.Code + Environment.NewLine);
                s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                s.Append("Message " + Exception.Message + Environment.NewLine);
                s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                File.AppendAllText("Log\\prospect.txt", s.ToString());
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\prospect.txt", sb.ToString());
                sb.Clear();
            }


        }
        public static void SendSalesDocument(string ctnum)
        {
            List<DocumentVente> documents = ControllerClient.GetClientSalesDocument(ctnum);
            File.AppendAllText("Log\\SendSales.txt",documents.Count().ToString());
            foreach (DocumentVente item in documents)
            {
                Boolean exist = false;
                try
                {
                    //deleteLigneDocument(item.NumPiece);
                    //exist = deleteDocument(item.NumPiece);
                }
                catch (ZCRMException Exception)
                {
                    StringBuilder s = new StringBuilder();
                    s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                    s.Append("Source " + Exception.Source + Environment.NewLine);
                    s.Append("HelpLink " + Exception.HelpLink + Environment.NewLine);
                    s.Append("Code " + Exception.Code + Environment.NewLine);
                    s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                    s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                    s.Append("Message " + Exception.Message + Environment.NewLine);
                    s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                    s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", s.ToString());
                }
                
                //File.AppendAllText("Log\\SendDocument.txt", item.Client.Length + Environment.NewLine);
                ZCRMRestClient restClient = ZCRMRestClient.GetInstance();
                ZCRMRecord IRecord = new ZCRMRecord("Vendors");
                ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
                IRecord.Owner = user;
                //IRecord.SetFieldValue("Account_Name", "["+item.Client+"]");
                if (!String.IsNullOrEmpty(item.DateCreation.ToString()))
                {
                    IRecord.SetFieldValue("Date_de_creation", item.DateCreation.ToString("yyyy-MM-dd"));
                }
                IRecord.SetFieldValue("Total_HT", item.Total_HT);
                IRecord.SetFieldValue("Vendor_Name", item.NumPiece);
                if (!String.IsNullOrEmpty(item.Reference))
                {
                    IRecord.SetFieldValue("Reference", item.Reference);
                }
                IRecord.SetFieldValue("Type_Document", item.DocumentType);
                if (item.Client.Length == 18)
                {
                    IRecord.SetFieldValue("Account_Name", Int64.Parse(item.Client));//245194000000989004);
                    File.AppendAllText("Log\\SendDocument.txt", "OK item" + Environment.NewLine);
                }
                else
                {
                    
                    var clientsSageObj = Singleton.SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(item.Client);
                    Client client = new Client(clientsSageObj);
                    SendClient(client);
                    var clientsSageObj2 = Singleton.SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(item.Client);
                    Client client2 = new Client(clientsSageObj2);

                    IRecord.SetFieldValue("Account_Name", Int64.Parse(client2.ZohoEntityId));
                    File.AppendAllText("Log\\SendDocument.txt", client2.ZohoEntityId.ToString() + Environment.NewLine);

                }
                List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
                try
                {
                    ZCRMModule module = restClient.GetModuleInstance("Vendors");
                    BulkAPIResponse<ZCRMRecord> response = module.CreateRecords(records);
                    string idDoc="";
                    string created = "";
                    foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
                    {
                        idDoc = eResponse.ResponseJSON["details"]["id"].ToString();
                        created = eResponse.Code.ToString();
                    }
                        if (created.Equals("SUCCESS"))
                        {
                            //create document vente ligne;
                            ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                            List<ZCRMRecord> records1 = new List<ZCRMRecord>();
                            //get document and fetch lines
                            IBODocumentVente3 doc = null;

                            if (item.DocumentType.Equals("Devis"))
                            {
                                doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteDevis, item.NumPiece);
                            }
                            if (item.DocumentType.Equals("Bon de commande"))
                            {
                                doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, item.NumPiece);
                            }
                            if (item.DocumentType.Equals("Bon de livraison"))
                            {
                                doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteLivraison, item.NumPiece);
                            }
                            if (item.DocumentType.Equals("Facture"))
                            {
                                doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFacture, item.NumPiece);
                            }
                            if (item.DocumentType.Equals("Facture comptabilisée"))
                            {
                                doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFactureCpta, item.NumPiece);
                            }
                            
                            File.AppendAllText("Log\\SendDocument.txt", doc.DO_Piece + Environment.NewLine);
                            foreach (IBODocumentVenteLigne3 ligne3 in doc.FactoryDocumentLigne.List)
                            {

                            //deleteLigneDocument(ligne3);
                                if (ligne3.Article == null)
                                {
                                    File.AppendAllText("Log\\SendDocumentError.txt", "erreur avec document  : "+ doc.DO_Piece + Environment.NewLine);
                                    continue;
                                }
                                ZCRMRecord IRecord1 = new ZCRMRecord("Solutions");
                                IRecord1.Owner = user;
                                IRecord1.SetFieldValue("Quantite", ligne3.DL_Qte.ToString());
                                IRecord1.SetFieldValue("Prix_unitaire_HT", ligne3.DL_MontantHT.ToString());
                                IRecord1.SetFieldValue("Designation", ligne3.Article.AR_Design);
                                IRecord1.SetFieldValue("Solution_Title", ligne3.Article.AR_Ref);
                                IRecord1.SetFieldValue("Num_Piece", Int64.Parse(idDoc));//ID of document entity
                                records1.Add(IRecord1);
                                File.AppendAllText("Log\\SendDocument.txt", "Trying to create this ligne : "+ligne3.Article.AR_Ref+" "+ligne3.DL_Qte.ToString() +" "+ ligne3.DL_MontantHT.ToString() +" "+ ligne3.Article.AR_Design + Environment.NewLine);
                            }
                            try
                            {
                                ZCRMModule module1 = restClient1.GetModuleInstance("Solutions");
                                BulkAPIResponse<ZCRMRecord> response1 = module1.CreateRecords(records1);
                            }
                            catch (ZCRMException Exception)
                            {

                                StringBuilder s = new StringBuilder();
                                s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                                s.Append("Code " + Exception.Code + Environment.NewLine);
                                s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                                s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                                s.Append("Message " + Exception.Message + Environment.NewLine);
                                s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                                s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                                File.AppendAllText("Log\\SendDocument.txt", s.ToString());
                            }
                            
                        }
                }
                catch (ZCRMException Exception)
                {

                    StringBuilder s = new StringBuilder();
                    s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                    s.Append("Code " + Exception.Code + Environment.NewLine);
                    s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                    s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                    s.Append("Message " + Exception.Message + Environment.NewLine);
                    s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                    s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", s.ToString());
                }
                catch (Exception s)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\doc.txt", sb.ToString());
                    sb.Clear();
                }

            }
            MessageBox.Show("end document sync", "ok",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);
        }
        public static void SendSalesDocument()
        {
            /*
             * à corriger
             * 
             * int test = GetDocuementRecordsToDelete();
            for (int i = 1; i < test; i++)
            {
                GetDocuementRecordsToDelete();
            }
            int line = deleteDocumentLine();
            for (int i = 1; i < line; i++)
            {
                deleteDocumentLine();
            }*/
            //File.AppendAllText("Log\\SendDocument.txt", "document number :" + test.Count  + Environment.NewLine);

            File.AppendAllText("Log\\SendDocument.txt", "Start sending all documents :" + Environment.NewLine);
            List<DocumentVente> documents = ControllerClient.GetClientsSalesDocument();
            File.AppendAllText("Log\\SendDocument.txt", "document number :"+documents.Count.ToString() + "List of all documents : " + Environment.NewLine);

            foreach (DocumentVente item in documents)
            {
                

            
            File.AppendAllText("Log\\SendDocument.txt", "Creating : "+item.NumPiece + Environment.NewLine);
            ZCRMRestClient restClient = ZCRMRestClient.GetInstance();
            ZCRMRecord IRecord = new ZCRMRecord("Vendors");
            ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
            IRecord.Owner = user;
            //IRecord.SetFieldValue("Account_Name", "["+item.Client+"]");
            if (!String.IsNullOrEmpty(item.DateCreation.ToString()))
            {
                IRecord.SetFieldValue("Date_de_creation", item.DateCreation.ToString("yyyy-MM-dd"));
            }
            IRecord.SetFieldValue("Total_HT", item.Total_HT);
            IRecord.SetFieldValue("Vendor_Name", item.NumPiece);
            if (!String.IsNullOrEmpty(item.Reference))
            {
                IRecord.SetFieldValue("Reference", item.Reference);
            }
            IRecord.SetFieldValue("Type_Document", item.DocumentType);
            if (item.Client.Length == 18)
            {
                IRecord.SetFieldValue("Account_Name", Int64.Parse(item.Client));//245194000000989004);
                //File.AppendAllText("Log\\SendDocument.txt", "OK item" + Environment.NewLine);
            }
            else
            {

                var clientsSageObj = Singleton.SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(item.Client);
                Client client = new Client(clientsSageObj);
                SendClient(client);
                var clientsSageObj2 = Singleton.SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(item.Client);
                Client client2 = new Client(clientsSageObj2);
                File.AppendAllText("Log\\SendDocument.txt", "Creating client in zoho"+client2.ZohoEntityId.ToString() + Environment.NewLine);
                IRecord.SetFieldValue("Account_Name", Int64.Parse(client2.ZohoEntityId));


            }
            List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
            try
            {
                ZCRMModule module = restClient.GetModuleInstance("Vendors");
                BulkAPIResponse<ZCRMRecord> response = module.CreateRecords(records);
                string idDoc = "";
                string created = "";
                foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
                {
                    idDoc = eResponse.ResponseJSON["details"]["id"].ToString();
                    created = eResponse.Code.ToString();
                }
                if (created.Equals("SUCCESS"))
                {

                    //create document vente ligne;
                    ZCRMRestClient restClient1 = ZCRMRestClient.GetInstance();
                    List<ZCRMRecord> records1 = new List<ZCRMRecord>();
                    //get document and fetch lines
                    IBODocumentVente3 doc = null;

                    if (item.DocumentType.Equals("Devis"))
                    {
                        doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteDevis, item.NumPiece);
                    }
                    if (item.DocumentType.Equals("Bon de commande"))
                    {
                        doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, item.NumPiece);
                    }
                    if (item.DocumentType.Equals("Bon de livraison"))
                    {
                        doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteLivraison, item.NumPiece);
                    }
                    if (item.DocumentType.Equals("Facture"))
                    {
                        doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFacture, item.NumPiece);
                    }
                    if (item.DocumentType.Equals("Facture comptabilisée"))
                    {
                        doc = SingletonConnection.Instance.Gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFactureCpta, item.NumPiece);
                    }
                    File.AppendAllText("Log\\SendDocument.txt", "Creating Document Ligne" + doc.DO_Piece + Environment.NewLine);

                    foreach (IBODocumentVenteLigne3 ligne3 in doc.FactoryDocumentLigne.List)
                    {
                        if (ligne3.Article == null)
                        {
                            File.AppendAllText("Log\\SendDocumentError.txt", "erreur avec document  : " + doc.DO_Piece + Environment.NewLine);
                            continue;
                        }
                        ZCRMRecord IRecord1 = new ZCRMRecord("Solutions");
                        IRecord1.Owner = user;
                        IRecord1.SetFieldValue("Quantite", ligne3.DL_Qte.ToString());
                        IRecord1.SetFieldValue("Prix_unitaire_HT", ligne3.DL_MontantHT.ToString());
                        IRecord1.SetFieldValue("Designation", ligne3.Article.AR_Design);
                        IRecord1.SetFieldValue("Solution_Title", ligne3.Article.AR_Ref);
                        IRecord1.SetFieldValue("Num_Piece", Int64.Parse(idDoc));//ID of document entity
                        records1.Add(IRecord1);
                        File.AppendAllText("Log\\SendDocument.txt", "Trying to create this ligne : " + ligne3.Article.AR_Ref + " " + ligne3.DL_Qte.ToString() + " " + ligne3.DL_MontantHT.ToString() + " " + ligne3.Article.AR_Design + Environment.NewLine);
                    }
                    try
                    {
                        ZCRMModule module1 = restClient1.GetModuleInstance("Solutions");
                        BulkAPIResponse<ZCRMRecord> response1 = module1.CreateRecords(records1);
                    }
                    catch (ZCRMException Exception)
                    {

                        StringBuilder s = new StringBuilder();
                        s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                        s.Append("Code " + Exception.Code + Environment.NewLine);
                        s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                        s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                        s.Append("Message " + Exception.Message + Environment.NewLine);
                        s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                        s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", s.ToString());
                    }

                }
            }
            catch (ZCRMException Exception)
            {

                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Code " + Exception.Code + Environment.NewLine);
                s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                s.Append("Message " + Exception.Message + Environment.NewLine);
                s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", s.ToString());
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\doc.txt", sb.ToString());
                sb.Clear();
            }
        }
        }
        public static int deleteDocumentLine()
        {
            List<long> toDelete = new List<long>();
            ZCRMModule moduleIns = ZCRMModule.GetInstance("Solutions"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            int page = 1;
            int perPage = 100;
            List<ZCRMRecord> records1;
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Solutions", Int64.Parse(Utils.UtilsConfig.CvIdDocumentLine), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        try
                        {
                            //deleteLigneDocument(Int64.Parse(record1.EntityId.ToString()));
                        }
                        catch (ZCRMException e)
                        {

                        }

                        //File.AppendAllText("Log\\SendDocument.txt", DateTime.Now + "  " + record1.EntityId.ToString() + Environment.NewLine);
                        toDelete.Add(Int64.Parse(record1.EntityId.ToString()));
                    }
                    page++;
                    BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
                    List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
                    foreach (EntityResponse eResponse1 in entityResponses)
                    {
                        File.AppendAllText("Log\\SendDocument.txt", "Delete ligne Docuement" + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
                    }
                    toDelete = null;
                } while (records1.Count == 100);

            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
                    List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
                    foreach (EntityResponse eResponse1 in entityResponses)
                    {
                        File.AppendAllText("Log\\SendDocument.txt", "Delete ligne Docuement" + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
                    }
                    return page;
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("Log\\SendDocument.txt", ex.InnerException.ToString() + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", ex.Message.ToString() + Environment.NewLine);
            }
            return page;
        }
        /*
        try
        {
            List<long> toDelete = new List<long>();
            ZCRMModule moduleIns = ZCRMModule.GetInstance("Vendors");
            BulkAPIResponse<ZCRMRecord> response = moduleIns.SearchByCriteria("(Vendor_Name:equals:" + docRef + ")");
            List<ZCRMRecord> records = response.BulkData;
            foreach (ZCRMRecord record in records)
            {
                foreach (var item in record.Data)
                {
                    File.AppendAllText("Log\\docData.txt", item.Key + ":" + item.Value);
                }
                toDelete.Add(Int64.Parse(record.EntityId.ToString()));
                deleteLigneDocument(Int64.Parse(record.EntityId.ToString()));
            }
            File.AppendAllText("Log\\SendDocument.txt", "Document count : " + toDelete.Count() + Environment.NewLine);
            BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
            List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
            foreach (EntityResponse eResponse1 in entityResponses)
            {
                File.AppendAllText("Log\\SendDocument.txt", "Delete Docuement" + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
            }
            return true;
        }
        catch (ZCRMException Exception)
        {
            StringBuilder s = new StringBuilder();
            s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
            s.Append("Source " + Exception.Source + Environment.NewLine);
            s.Append("HelpLink " + Exception.HelpLink + Environment.NewLine);
            s.Append("Code " + Exception.Code + Environment.NewLine);
            s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
            s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
            s.Append("Message " + Exception.Message + Environment.NewLine);
            s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
            s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
            File.AppendAllText("Log\\SendDocument.txt", s.ToString());
            return false;
        }
}*/
        public static void deleteLigneDocument(long entityID)//IBODocumentVenteLigne3 ligne)
        {
            try
            {

                List<long> toDelete = new List<long>();
                ZCRMModule moduleIns = ZCRMModule.GetInstance("Solutions");
                ZCRMRecord recordIns = ZCRMRecord.GetInstance("Vendors", entityID); //module api name with record id
                BulkAPIResponse<ZCRMRecord> response = recordIns.GetRelatedListRecords("Lignes_Document_de_Vente"); //RelatedList api name
                List<ZCRMRecord> relatedRecordsLists = response.BulkData;
                //BulkAPIResponse<ZCRMRecord> response = moduleIns.SearchByCriteria("((Quantite:equals:" + ligne.DL_Qte + ")and(Solution_Title:equals:" + ligne.Article.AR_Ref + "))");
                //List<ZCRMRecord> records = response.BulkData;
                foreach (ZCRMRecord record in relatedRecordsLists)
                {
                    //record.Data["Num_Piece"]["name"]
                    foreach (var item in record.Data)
                    {
                        File.AppendAllText("Log\\docLigneData.txt", item.Key + ":" + item.Value + Environment.NewLine);
                    }

                    toDelete.Add(Int64.Parse(record.EntityId.ToString()));
                }
                BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
                List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
                foreach (EntityResponse eResponse1 in entityResponses)
                {
                    File.AppendAllText("Log\\SendDocument.txt", "Delete ligne Docuement" + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                    File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
                }
            }
            catch (ZCRMException Exception)
            {
                StringBuilder s = new StringBuilder();
                s.Append(DateTime.Now + Exception.HttpStatusCode.ToString() + Environment.NewLine);
                s.Append("Source " + Exception.Source + Environment.NewLine);
                s.Append("HelpLink " + Exception.HelpLink + Environment.NewLine);
                s.Append("Code " + Exception.Code + Environment.NewLine);
                s.Append("IsAPIException " + Exception.IsAPIException + Environment.NewLine);
                s.Append("IsSDKException " + Exception.IsSDKException + Environment.NewLine);
                s.Append("Message " + Exception.Message + Environment.NewLine);
                s.Append("ErrorDetails " + Exception.ErrorDetails + Environment.NewLine);
                s.Append("ErrorMsg " + Exception.ErrorMsg + Environment.NewLine + Environment.NewLine + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", s.ToString());
            }
        }
        public static void SendArticle(Article article)
        {
            ZCRMRestClient restClient = ZCRMRestClient.GetInstance();
            ZCRMRecord IRecord = new ZCRMRecord("Products");
            ZCRMUser user = ZCRMUser.GetInstance(Int64.Parse(ConfigurationManager.AppSettings["IdUserAPI"]));//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
            IRecord.Owner = user;
            IRecord.SetFieldValue("Product_Name", article.Designation );
            IRecord.SetFieldValue("Product_Code", article.Reference);
            IRecord.SetFieldValue("Product_Active", !article.Sommeil);
            IRecord.SetFieldValue("Qty_in_Stock", article.Stock);
            IRecord.SetFieldValue("Description", article.Designation);
            List<ZCRMRecord> records = new List<ZCRMRecord> { IRecord };
            ZCRMModule module = restClient.GetModuleInstance("Products");
            BulkAPIResponse<ZCRMRecord> response = module.CreateRecords(records);
            foreach (EntityResponse eResponse in response.BulkEntitiesResponse)
            {
                Console.WriteLine(eResponse.ResponseJSON);//Fetches response as JSON
                Console.WriteLine(eResponse.Status);//Fetches status value present in the response
                Console.WriteLine(eResponse.Code);//Fetches code value present in the response
                Console.WriteLine(eResponse.Message);//Fetches message value present in the response
            }
        }
        private static BulkAPIResponse<ZCRMRecord> GetModuleData(string ModuleName,Int64 CvId,int page, int perPage,List<string> FieldList)
        {
            List<string> ct_num = new List<string>();
            ZCRMModule module1 = ZCRMModule.GetInstance(ModuleName); //GetModuleInstance(ApiName.Key);
            BulkAPIResponse<ZCRMRecord> response1 = module1.GetRecords(CvId, page, perPage, FieldList);
            return response1;
        }
        public static List<string> GetListOfprospectToProcessZoho()
        {
            List<string> ct_num = new List<string>(); 
            ZCRMModule module1 = ZCRMModule.GetInstance("Leads"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            fieldsCRM.Add("CT_Num");
            fieldsCRM.Add("Id");
            int page = 1;
            int perPage = 200;
            List<ZCRMRecord> records1;
            SingletonUI.Instance.ProspectList = new List<CtnumId>();
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Leads", Int64.Parse(Utils.UtilsConfig.CvIdProspect), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        ct_num.Add(record1.Data["CT_Num"].ToString());
                        SingletonUI.Instance.ProspectList.Add(new CtnumId(record1.Data["CT_Num"].ToString(), Int64.Parse(record1.EntityId.ToString())));
                    }
                    page++;
                } while (records1.Count == 200);
            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    return ct_num;
                }
            }
            return ct_num;
        }
        public static List<string> GetListOfClientToProcessZoho()
        {
            List<string> ct_num = new List<string>();
            ZCRMModule module1 = ZCRMModule.GetInstance("Accounts"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            fieldsCRM.Add("CT_Num");
            fieldsCRM.Add("id");
            int page = 1;
            int perPage = 200;
            List<ZCRMRecord> records1;
            SingletonUI.Instance.AccountList = new List<CtnumId>();
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Accounts", Int64.Parse(Utils.UtilsConfig.CvIdAccounts), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        ct_num.Add(record1.Data["CT_Num"].ToString());
                        SingletonUI.Instance.AccountList.Add(new CtnumId(record1.Data["CT_Num"].ToString(), Int64.Parse(record1.Data["id"].ToString())));
                    }
                    page++;
                } while (records1.Count == 200);
            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    return ct_num;
                }
            }
            return ct_num;
        }
        public static List<ZCRMRecord> GetRecordsOfprospectToProcessZoho()
        {
            List<ZCRMRecord> Prospects = new List<ZCRMRecord>();
            ZCRMModule module1 = ZCRMModule.GetInstance("Leads"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            fieldsCRM.Add("Company");
            fieldsCRM.Add("Last_Name");
            fieldsCRM.Add("First_Name");
            fieldsCRM.Add("Phone");
            fieldsCRM.Add("CT_Num");
            fieldsCRM.Add("Facebook");
            fieldsCRM.Add("LinkedIn");
            fieldsCRM.Add("Interlocuteur");
            fieldsCRM.Add("Identifiant_TVA");
            fieldsCRM.Add("Sommeil");
            fieldsCRM.Add("Commentaire");
            fieldsCRM.Add("Groupe_Tarifaire");
            fieldsCRM.Add("Fax");
            fieldsCRM.Add("Nom_Contact");
            fieldsCRM.Add("Prenom_Contact");
            fieldsCRM.Add("Telecopie_Contact");
            fieldsCRM.Add("Telephone_Contact");
            fieldsCRM.Add("Portable_Contact");
            fieldsCRM.Add("Civilite_Contact");
            fieldsCRM.Add("Skype_Contact");
            fieldsCRM.Add("Facebook_Contact");
            fieldsCRM.Add("Fonction_Contact");
            fieldsCRM.Add("Service_Contact");
            fieldsCRM.Add("Email_Contact");
            fieldsCRM.Add("LinkedIn_Contact");
            fieldsCRM.Add("Email");
            fieldsCRM.Add("Street");
            fieldsCRM.Add("City");
            fieldsCRM.Add("Zip_Code");
            fieldsCRM.Add("Country");
            fieldsCRM.Add("SIRET");
            fieldsCRM.Add("Centrale_d_achat");
            fieldsCRM.Add("Representant");
            fieldsCRM.Add("Code_NAF_Sage");
            fieldsCRM.Add("Qualite");
            fieldsCRM.Add("Groupe");
            fieldsCRM.Add("Enseigne");
            int page = 1;
            int perPage = 200;
            List<ZCRMRecord> records1;
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Leads", Int64.Parse(Utils.UtilsConfig.CvIdProspect), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        if (String.IsNullOrEmpty(record1.Data["CT_Num"].ToString()))
                        {
                            Prospects.Add(record1);
                        }
                    }
                    page++;
                } while (records1.Count == 200);
            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    return Prospects;
                }
            }
            return Prospects;
        }

        public static void TransformToCLient(string ctnum,string id)
        {
            IBOClient3 prospect = SingletonConnection.Instance.Compta.FactoryProspect.ReadNumero(ctnum);
            IPMConversionClient conversionClient = SingletonConnection.Instance.Gescom.CreateProcess_ConversionClient(prospect);
            conversionClient.Intitule = prospect.CT_Intitule;
            conversionClient.NumPayeur = prospect;
            try
            {
                conversionClient.Process();
                try
                {
                    IBOClient3 clientSage = SingletonConnection.Instance.Compta.FactoryClient.ReadNumero(ctnum);
                    var infolibreField = Singleton.SingletonConnection.Instance.Compta.FactoryTiers.InfoLibreFields;
                    int compteur = 1;

                    foreach (var infoLibreValue in clientSage.InfoLibre)
                    {
                        if (infolibreField[compteur].Name.Equals("ZohoEntityID"))
                        {
                            clientSage.InfoLibre[compteur] = id;
                            break;
                        }
                        compteur++;
                    }
                    //clientSage.InfoLibre['6'] = eResponse.ResponseJSON["details"]["id"].ToString();
                    clientSage.Write();
                }
                catch (Exception s)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\zohoEntityId.txt", sb.ToString());
                    sb.Clear();
                }
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + "Erreur transformation de prospect en client : " + ctnum + Environment.NewLine);
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\TransformProspect.txt", sb.ToString());
                sb.Clear();
            }
        }
        public static List<ZCRMRecord> GetRecordsOfAccountsToProcessZoho()
        {
            List<ZCRMRecord> Prospects = new List<ZCRMRecord>();
            ZCRMModule module1 = ZCRMModule.GetInstance("Accounts"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            fieldsCRM.Add("CT_Num");
            fieldsCRM.Add("id");

            int page = 1;
            int perPage = 200;
            List<ZCRMRecord> records1;
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Accounts", Int64.Parse(Utils.UtilsConfig.CvIdAccounts), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        if (SingletonConnection.Instance.Compta.FactoryProspect.ExistNumero(record1.Data["CT_Num"].ToString()))
                        {
                            Prospects.Add(record1);
                        }
                    }
                    page++;
                } while (records1.Count == 200);
            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    return Prospects;
                }
            }
            return Prospects;
        }
        public static int GetDocuementRecordsToDelete()
        {
            List<long> toDelete = new List<long>();
            ZCRMModule moduleIns = ZCRMModule.GetInstance("Vendors"); //GetModuleInstance(ApiName.Key);
            List<string> fieldsCRM = new List<string>();
            int page = 1;
            int perPage = 100;
            List<ZCRMRecord> records1;
            try
            {
                do
                {
                    BulkAPIResponse<ZCRMRecord> response1 = GetModuleData("Vendors", Int64.Parse(Utils.UtilsConfig.CvIdDocument), page, perPage, fieldsCRM);
                    records1 = null;
                    records1 = response1.BulkData;

                    foreach (ZCRMRecord record1 in records1)
                    {
                        try
                        {
                            //deleteLigneDocument(Int64.Parse(record1.EntityId.ToString()));
                        }
                        catch (ZCRMException e )
                        {

                        }
                        
                        //File.AppendAllText("Log\\SendDocument.txt", DateTime.Now + "  " + record1.EntityId.ToString() + Environment.NewLine);
                        toDelete.Add(Int64.Parse(record1.EntityId.ToString()));
                    }
                    page++;
                    BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
                    List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
                    foreach (EntityResponse eResponse1 in entityResponses)
                    {
                        File.AppendAllText("Log\\SendDocument.txt", "Delete ligne Docuement" + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
                    }
                    toDelete = null;
                } while (records1.Count == 100);
                
            }
            catch (ZCRMException ex)
            {
                if (ex.Code.ToString().Equals("NO_CONTENT"))
                {
                    BulkAPIResponse<ZCRMEntity> response2 = moduleIns.DeleteRecords(toDelete);
                    List<EntityResponse> entityResponses = response2.BulkEntitiesResponse;
                    foreach (EntityResponse eResponse1 in entityResponses)
                    {
                        File.AppendAllText("Log\\SendDocument.txt", "Delete ligne Docuement" + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Code.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Message.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.Data.ToString() + Environment.NewLine);
                        File.AppendAllText("Log\\SendDocument.txt", eResponse1.ResponseJSON.ToString() + Environment.NewLine);
                    }
                    return page;
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("Log\\SendDocument.txt", ex.InnerException.ToString() + Environment.NewLine);
                File.AppendAllText("Log\\SendDocument.txt", ex.Message.ToString() + Environment.NewLine);
            }
            return page;
        }
        public ZohoController()
        {
            
            //var leadinfo = new { "Company":input.Company_Name,"Last_Name":input.Last_Name,"Phone":input.Phone,"Email":input.Email_Address,"PO_Box":input.PO_Box,"Country":input.Country};
            /* Inert Record */
            /*
            ZCRMRestClient restClient = ZCRMRestClient.GetInstance();
            ZCRMRecord IRecord = new ZCRMRecord("Accounts");
            ZCRMUser user = ZCRMUser.GetInstance(20069463844);//ZCRMUser.GetInstance(((ZCRMUser)userResponse.Data).Id);USER WHO CREATED THE API TOKEN
            IRecord.Owner = user;
            IRecord.SetFieldValue("Account_Name", "Test insert");
            List<ZCRMRecord> records1 = new List<ZCRMRecord> { IRecord };
            ZCRMModule module = restClient.GetModuleInstance("Accounts");
            BulkAPIResponse<ZCRMRecord> response1 = module.CreateRecords(records1);
            foreach (EntityResponse eResponse in response1.BulkEntitiesResponse)
            {
                Console.WriteLine(eResponse.ResponseJSON);//Fetches response as JSON
                Console.WriteLine(eResponse.Status);//Fetches status value present in the response
                Console.WriteLine(eResponse.Code);//Fetches code value present in the response
                Console.WriteLine(eResponse.Message);//Fetches message value present in the response
            }

            /* Get Module Data */
            /*
            ZCRMModule module = ZCRMModule.GetInstance("DocumentsVentes"); //GetModuleInstance(ApiName.Key);
            BulkAPIResponse<ZCRMRecord> response = module.GetRecords();
            List<ZCRMRecord> records = response.BulkData;
            foreach (ZCRMRecord record in records)
            {
                Console.WriteLine(record.EntityId);//Fetches record id
                foreach (var item in record.Data)
                {
                    Console.WriteLine(item.Key + ":" + item.Value);
                }
                Console.WriteLine(record.Data);//Fetches a dictionary which has field api name and it’s value as key and value in that dictionary 
                Console.WriteLine(record.CreatedBy);//Fetches created by of the record
                Console.WriteLine(record.ModuleAPIName);//Fetches record’s module api name
            }*/
            /*ZCRMModule moduleIns = ZCRMModule.GetInstance("Accounts"); //module api name
            BulkAPIResponse<ZCRMModuleRelation> responseL = moduleIns.GetRelatedLists();
            List<ZCRMModuleRelation> relatedLists = responseL.BulkData;

            ZCRMRecord recordIns = ZCRMRecord.GetInstance("Accounts", 245194000000876044); //module api name with record id
            BulkAPIResponse<ZCRMRecord> responseList = recordIns.GetRelatedListRecords("DocumentsVentes"); //RelatedList api name
            List<ZCRMRecord> relatedRecordsLists = responseList.BulkData;
            */

            ZCRMModule module1 = ZCRMModule.GetInstance("Accounts"); //GetModuleInstance(ApiName.Key);
            BulkAPIResponse<ZCRMRecord> response1 = module1.GetRecords();
            List<ZCRMRecord> records1 = response1.BulkData;
            foreach (ZCRMRecord record1 in records1)
            {
                Console.WriteLine(record1.EntityId);//Fetches record id
                foreach (var item in record1.Data)
                {
                    Console.WriteLine(item.Key + ":" + item.Value);
                }
                Console.WriteLine(record1.Data);//Fetches a dictionary which has field api name and it’s value as key and value in that dictionary 
                Console.WriteLine(record1.CreatedBy);//Fetches created by of the record
                Console.WriteLine(record1.ModuleAPIName);//Fetches record’s module api name
            }
            Console.WriteLine("test");
            /** Get Module Fields */
            /*
            
            ZCRMModule moduleIns = ZCRMModule.GetInstance("Accounts"); //module api name
                BulkAPIResponse<ZCRMField> response = moduleIns.GetAllFields();
                Console.WriteLine(response.HttpStatusCode);
                List<ZCRMField> fields = response.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin ACCOUNTS FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field in fields)
            {
                Console.WriteLine(field.DisplayName);
                Console.WriteLine(field.DataType);
                Console.WriteLine(field.Mandatory);
                Console.WriteLine(field.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END ACCOUNTS FIELD");
            Console.WriteLine("************");
            /*
            ZCRMModule moduleIns1 = ZCRMModule.GetInstance("Contacts"); //module api name
            BulkAPIResponse<ZCRMField> response1 = moduleIns1.GetAllFields();
            Console.WriteLine(response1.HttpStatusCode);
            List<ZCRMField> fields1 = response1.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin CONTACTS FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field1 in fields1)
            {
                Console.WriteLine(field1.DisplayName);
                Console.WriteLine(field1.DataType);
                Console.WriteLine(field1.Mandatory);
                Console.WriteLine(field1.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END  CONTACTS FIELD");
            Console.WriteLine("************");

            ZCRMModule moduleIns2 = ZCRMModule.GetInstance("Products"); //module api name
            BulkAPIResponse<ZCRMField> response2 = moduleIns2.GetAllFields();
            Console.WriteLine(response2.HttpStatusCode);
            List<ZCRMField> fields2 = response2.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin PRODUCT FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field2 in fields2)
            {
                Console.WriteLine(field2.DisplayName);
                Console.WriteLine(field2.DataType);
                Console.WriteLine(field2.Mandatory);
                Console.WriteLine(field2.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END PRODUCT FIELD");
            Console.WriteLine("************");

            ZCRMModule moduleIns3 = ZCRMModule.GetInstance("Quotes"); //module api name
            BulkAPIResponse<ZCRMField> response3 = moduleIns3.GetAllFields();
            Console.WriteLine(response3.HttpStatusCode);
            List<ZCRMField> fields3 = response3.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin DEVIS FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field2 in fields3)
            {
                Console.WriteLine(field2.DisplayName);
                Console.WriteLine(field2.DataType);
                Console.WriteLine(field2.Mandatory);
                Console.WriteLine(field2.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END DEVIS FIELD");
            Console.WriteLine("************");

            ZCRMModule moduleIns4 = ZCRMModule.GetInstance("Invoices"); //module api name
            BulkAPIResponse<ZCRMField> response4 = moduleIns4.GetAllFields();
            Console.WriteLine(response4.HttpStatusCode);
            List<ZCRMField> fields4 = response4.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin Invoices FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field2 in fields4)
            {
                Console.WriteLine(field2.DisplayName);
                Console.WriteLine(field2.DataType);
                Console.WriteLine(field2.Mandatory);
                Console.WriteLine(field2.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END Invoices FIELD");
            Console.WriteLine("************");

            ZCRMModule moduleIns5 = ZCRMModule.GetInstance("Purchase_Orders"); //module api name
            BulkAPIResponse<ZCRMField> response5 = moduleIns5.GetAllFields();
            Console.WriteLine(response5.HttpStatusCode);
            List<ZCRMField> fields5 = response5.BulkData; //fields - list of ZCRMField instance
            Console.WriteLine("begin Bons de commande FIELD");
            Console.WriteLine("************");
            foreach (ZCRMField field2 in fields5)
            {
                Console.WriteLine(field2.DisplayName);
                Console.WriteLine(field2.DataType);
                Console.WriteLine(field2.Mandatory);
                Console.WriteLine(field2.ApiName);

                Console.WriteLine("************");
            }
            Console.WriteLine("END Bons de commande FIELD");
            Console.WriteLine("************");
            */
            /*Purchase_Orders
                Console.WriteLine(field.CreatedSource);
                Console.WriteLine(field.FormulaReturnType);
                Console.WriteLine(field.JsonType);
                Console.WriteLine(field.ToolTip);
                Console.WriteLine(field.CustomField);
                Console.WriteLine(field.DefaultValue);
                Dictionary<string, object> LookupDetails = field.LookupDetails;
                foreach (KeyValuePair<string, object> LookupDetail in LookupDetails)
                {
                    Console.WriteLine(LookupDetail.Key + ":" + LookupDetail.Value);
                }
                Console.WriteLine(field.Mandatory);
                Console.WriteLine(field.MaxLength);

                Dictionary<string, object> MultiselectLookup = field.MultiselectLookup;
                foreach (KeyValuePair<string, object> Lookup in MultiselectLookup)
                {
                    Console.WriteLine(Lookup.Key + ":" + Lookup.Value);
                }
                List<ZCRMPickListValue> pickListValues = field.PickListValues;
                foreach (ZCRMPickListValue pickListValue in pickListValues)
                {
                    Console.WriteLine(pickListValue.ActualName);
                    Console.WriteLine(pickListValue.DisplayName);
                }
                Console.WriteLine(field.Precision);
                Console.WriteLine(field.ReadOnly);
                Console.WriteLine(field.SequenceNo);

                Dictionary<string, object> SubformDetails = field.SubformDetails;
                foreach (KeyValuePair<string, object> SubformDetail in SubformDetails)
                {
                    Console.WriteLine(SubformDetail.Key + ":" + SubformDetail.Value);
                }
                Console.WriteLine(field.SubFormTabId);
                Console.WriteLine(field.Visible);
                Console.WriteLine(field.Webhook);
            }
        /*
        List<ZCRMRecord> listRecord = new List<ZCRMRecord>();
        ZCRMRecord record;
        record = new ZCRMRecord("Accounts");
        record.SetFieldValue("Account_Name", "Omar Magherbi");
        ZCRMUser owner = ZCRMUser.GetInstance(20069463844);//User Id
        record.Owner = owner;
        listRecord.Add(record);
        ZCRMModule moduleIns = ZCRMModule.GetInstance("Accounts"); //To get the Module instance
        BulkAPIResponse<ZCRMRecord> responseIns = moduleIns.UpdateRecords(listRecord); //To call the Update record method
        Console.WriteLine("HTTP Status Code:" + responseIns.HttpStatusCode);*/



        }
    }
}
