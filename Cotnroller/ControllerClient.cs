using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Objets100cLib;
using WebservicesSage.Object;
using WebservicesSage.Singleton;
using WebservicesSage.Utils;
using WebservicesSage.Utils.Enums;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Timers;
using ZCRMSDK.CRM.Library.CRUD;

namespace WebservicesSage.Cotnroller
{
    public static class ControllerClient
    {
        public static void LaunchService()
        {
            /*
            System.Timers.Timer timerUpdateStatut = new System.Timers.Timer();
            timerUpdateStatut.Elapsed += new ElapsedEventHandler(SendProspectCron);
            timerUpdateStatut.Interval = 30000;
            timerUpdateStatut.Enabled = true;
            //timer for the restart
            /*
            System.Timers.Timer timerRestartApplication = new System.Timers.Timer();
            timerRestartApplication.Elapsed += new ElapsedEventHandler(TimeToRestart);
            timerRestartApplication.Interval = 20000;
            timerRestartApplication.Enabled = true;*/
        }

        private static void TimeToRestart(object sender, ElapsedEventArgs e)
        {
            string[] date = UtilsConfig.CronTaskRestart.ToString().Split(':');
            int hours, minutes;
            hours = Int32.Parse(date[0]);
            minutes = Int32.Parse(date[1]);
            DateTime now = DateTime.Now;
            DateTime NowNoSecond = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, 0);
            DateTime firstRun = new DateTime(now.Year, now.Month, now.Day, hours, minutes, 0, 0);
            if (NowNoSecond == firstRun)
            {
                try
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("date de redémarrage : " + DateTime.Now + Environment.NewLine);
                    System.IO.File.AppendAllText("Log\\restart.txt", sb.ToString());
                    sb.Clear();
                    System.Diagnostics.Process.Start(Application.ExecutablePath);
                    Application.ExitThread();
                    Application.Exit();
                    Environment.Exit(200);
                }
                catch (Exception ex)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + ex.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + ex.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\restart.txt", sb.ToString());
                    sb.Clear();
                }

            }
            else
            {
                Console.WriteLine("do not restart");
            }
            
        }

        public static List<DocumentVente> GetClientSalesDocument(string dopiece)
        {
            List<DocumentVente> documents = new List<DocumentVente>();
            try
            {
                var gescom = SingletonConnection.Instance.Gescom;
                
                foreach (IBODocumentVente3 documentVente in gescom.FactoryDocumentVente.List)
                {
                    //File.AppendAllText("Log\\Sales.txt", documentVente.Client.CT_Num.ToString() +" "+ documentVente.DO_Type + Environment.NewLine);
                    if (documentVente.DO_Piece.Equals(dopiece))
                    {
                        var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                        var clientsSageObj = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                        Client client = new Client(clientsSageObj);
                        //string test = documentVente.DO_Piece + documentVente.Client.CT_Intitule + documentVente.DateCreation.ToShortDateString();
                        switch (documentVente.DO_Type)
                        {
                            case DocumentType.DocumentTypeVenteDevis:
                                DocumentVente document = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Devis", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document);
                                break;
                            case DocumentType.DocumentTypeVenteCommande:
                                DocumentVente document1 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Bon de commande", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document1);
                                break;
                            case DocumentType.DocumentTypeVenteLivraison:
                                DocumentVente document2 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Bon de livraison", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document2);
                                break;

                            case DocumentType.DocumentTypeVenteFacture:
                                DocumentVente document3 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Facture", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document3);
                                break;
                            case DocumentType.DocumentTypeVenteFactureCpta:
                                DocumentVente document4 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Facture comptabilisée", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document4);
                                break;
                            default:
                                break;

                        }
                    }
                }
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\GetSalesDocument.txt", sb.ToString());
                sb.Clear();
            }
            return documents;
        }
        public static List<DocumentVente> GetClientsSalesDocument()
        {
            List<DocumentVente> documents = new List<DocumentVente>();
            try
            {
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                var gescom = SingletonConnection.Instance.Gescom;
                foreach (IBODocumentVente3 documentVente in gescom.FactoryDocumentVente.List)
                {

                    switch (documentVente.DO_Type)
                    {
                        case DocumentType.DocumentTypeVenteDevis:
                            if (documentVente.Client.CT_Prospect)
                            {
                                break;
                            }
                            else
                            {
                                var clientsSageObj = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                                Client client = new Client(clientsSageObj);
                                if (!String.IsNullOrEmpty(client.ZohoEntityId))
                                {
                                    DocumentVente document = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client.ZohoEntityId, "Devis", documentVente.DateCreation, documentVente.DO_TotalHT);
                                    documents.Add(document);
                                }
                                else
                                {
                                    DocumentVente document = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, documentVente.Client.CT_Num, "Devis", documentVente.DateCreation, documentVente.DO_TotalHT);
                                    documents.Add(document);
                                }
                                break;
                            }

                        case DocumentType.DocumentTypeVenteCommande:
                            var clientsSageObj1 = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                            Client client1 = new Client(clientsSageObj1);
                            if (!String.IsNullOrEmpty(client1.ZohoEntityId))
                            {
                                DocumentVente document1 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client1.ZohoEntityId, "Bon de commande", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document1);
                            }
                            else
                            {
                                DocumentVente document1 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, documentVente.Client.CT_Num, "Bon de commande", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document1);
                            }
                            break;
                        case DocumentType.DocumentTypeVenteLivraison:
                            var clientsSageObj2 = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                            Client client2 = new Client(clientsSageObj2);
                            if (!String.IsNullOrEmpty(client2.ZohoEntityId))
                            {
                                DocumentVente document2 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client2.ZohoEntityId, "Bon de livraison", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document2);
                            }
                            else
                            {
                                DocumentVente document2 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, documentVente.Client.CT_Num, "Bon de livraison", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document2);
                            }
                            break;

                        case DocumentType.DocumentTypeVenteFacture:
                            var clientsSageObj3 = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                            Client client3 = new Client(clientsSageObj3);

                            if (!String.IsNullOrEmpty(client3.ZohoEntityId))
                            {
                                DocumentVente document3 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client3.ZohoEntityId, "Facture", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document3);
                            }
                            else
                            {
                                DocumentVente document3 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, documentVente.Client.CT_Num, "Facture", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document3);
                            }
                            break;
                        case DocumentType.DocumentTypeVenteFactureCpta:
                            var clientsSageObj4 = compta.FactoryClient.ReadNumero(documentVente.Client.CT_Num);
                            Client client4 = new Client(clientsSageObj4);

                            if (!String.IsNullOrEmpty(client4.ZohoEntityId))
                            {
                                DocumentVente document4 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, client4.ZohoEntityId, "Facture comptabilisée", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document4);
                            }
                            else
                            {
                                DocumentVente document4 = new DocumentVente(documentVente.DO_Piece, "EUR", documentVente.DO_Ref, documentVente.Client.CT_Num, "Facture comptabilisée", documentVente.DateCreation, documentVente.DO_TotalHT);
                                documents.Add(document4);
                            }
                            break;
                        default:
                            break;

                    }
                }
                
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\GetSalesDocument.txt", sb.ToString());
                sb.Clear();
            }
            return documents;
        }

        /// <summary>
        /// Permets de remonter toute la base clients de SAGE vers Prestashop
        /// Ne remonte que les clients avec un mail 
        /// </summary>
        public static void SendAllClients()
        {
            File.AppendAllText("Log\\AllClientTest.txt", "Worked!!!!2" + Environment.NewLine);
            var compta = SingletonConnection.Instance.Gescom.CptaApplication;
            var clientsSageObj = compta.FactoryClient.List;

            var clients = GetListOfClientToProcess(clientsSageObj);

            int increm = 100 / clients.Count;
            File.AppendAllText("Log\\AllClientTest.txt", clients.Count.ToString() + Environment.NewLine);
            foreach (Client client in clients)
            {
                
                if (UtilsConfig.ActiveClient.Equals("FALSE"))
                {
                    File.AppendAllText("Log\\AllClientTest.txt", "Worked!!!!5" + Environment.NewLine);
                    if (client.Sommeil)
                    {
                        File.AppendAllText("Log\\AllClientTest.txt", "Worked!!!!6" + Environment.NewLine);
                        continue;
                    }

                }

                    if (String.IsNullOrEmpty(client.ZohoEntityId) )
                    {
                    File.AppendAllText("Log\\AllClientTest.txt", "Inserted CLient : "+client.CT_NUM + Environment.NewLine);
                    ZohoController.SendClient(client);
                    }
                    else
                    {
                    File.AppendAllText("Log\\AllClientTest.txt", "Update CLient : " + client.CT_NUM + Environment.NewLine);
                    ZohoController.UpdateClient(client);
                    }

                // SingletonUI.Instance.ClientCircleProgress.Invoke((MethodInvoker)(() => SingletonUI.Instance.ClientCircleProgress.Value += increm));
                //SingletonUI.Instance.LogBox.Invoke((MethodInvoker)(() => SingletonUI.Instance.LogBox.AppendText("-- Processing Client -- " + client.Intitule + " END -- " + Environment.NewLine)));
            }

            MessageBox.Show("end client sync", "ok",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Information);

        }


        public static void SendClient(string ct_num)
        {
            try
            {
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                var clientsSageObj = compta.FactoryClient.ReadNumero(ct_num);
                var clients = GetClientToProcess(clientsSageObj);

                //int increm = 100 / clients.Count;
                File.AppendAllText("Log\\ClientTest.txt", clients.Count.ToString());
                foreach (Client client in clients)
                {
                    if (String.IsNullOrEmpty(client.ZohoEntityId))
                    {
                        ZohoController.SendClient(client);
                    }
                    else
                    {
                        ZohoController.UpdateClient(client);
                    }
                }
                var clientsSageObj1 = compta.FactoryClient.ReadNumero(ct_num);
                Client clients1 = new Client(clientsSageObj1);
                MessageBox.Show("end client sync"+clients1.ZohoEntityId, "ok",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);
            }
        }
        public static void SendProspect(string ct_num)
        {
            try
            {
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                var clientsSageObj = compta.FactoryProspect.ReadNumero(ct_num);
                var clients = GetProspectToProcess(clientsSageObj);

                //int increm = 100 / clients.Count;

                foreach (Prospect client in clients)
                {
                    if (String.IsNullOrEmpty(client.ZohoEntityId))
                    {
                        ZohoController.SendProspect(client);
                    }
                    else
                    {
                        ZohoController.UpdateProspect(client);
                    }
                }

                MessageBox.Show("end client sync", "ok",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);
            }
        }
        public static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address.ToUpper().Equals(email.ToUpper());
            }
            catch
            {
                return false;
            }
        }
        public static bool validTelephoneNo(string telNo)
        {
            return Regex.Match(telNo, @"\(?([0-9]{2})\)?([.-]?)([0-9]{2})\2([0-9]{2})\2([0-9]{2})\2([0-9]{2})").Success;
        }
        public static void SendProspect()
        {
            try
            {
                //ZohoController.CreateProspect();
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                var clientsSageObj = compta.FactoryProspect.List;
                var clients = GetListOfprospectToProcess(clientsSageObj);

                //int increm = 100 / clients.Count;
                File.AppendAllText("Log\\SendProspect.txt", "will create  : " + clients.Count.ToString() +" Prospect" + Environment.NewLine);
                foreach (Prospect client in clients)
                {
                    if (String.IsNullOrEmpty(client.ZohoEntityId))
                    {
                        if (!String.IsNullOrEmpty(client.Email))
                        {
                            if (!IsValidEmail(client.Email))
                            {
                                File.AppendAllText("Log\\SendProspectError.txt", "adresse mail invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                continue;
                            }
                        }
                        if (!String.IsNullOrEmpty(client.telephone))
                        {
                            if (!validTelephoneNo(client.telephone))
                            {
                                File.AppendAllText("Log\\SendProspectError.txt", "Telephone invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                continue;
                            }
                        }
                        if (!String.IsNullOrEmpty(client.telecopie))
                        {
                            if (!validTelephoneNo(client.telecopie))
                            {
                                File.AppendAllText("Log\\SendProspectError.txt", "Telecopie invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                continue;
                            }
                        }
                        if (!String.IsNullOrEmpty(client.Website))
                        {
                            if (IsValidEmail(client.Website))
                            {
                                File.AppendAllText("Log\\SendProspectError.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                continue;
                            }
                        }
                        if (!String.IsNullOrEmpty(client.Website))
                        {
                            if (validTelephoneNo(client.Website))
                            {
                                File.AppendAllText("Log\\SendProspectError.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                continue;
                            }
                        }
                        //string articleXML = UtilsSerialize.SerializeObject(client);
                        File.AppendAllText("Log\\SendProspect.txt", "try to send : " + client.Intitule.ToString() + Environment.NewLine);
                        ZohoController.SendProspect(client);
                        File.AppendAllText("Log\\SendProspect.txt", "Done : "+ client.CT_NUM.ToString()+ Environment.NewLine);
                    }
                    else
                    {
                        ZohoController.UpdateProspect(client);
                    }
                }
                //ZohoController.CreateProspect();

                MessageBox.Show("end lead sync", "ok",
                                   MessageBoxButtons.OK,
                                   MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + e.Message + Environment.NewLine);
                sb.Append(DateTime.Now + e.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\Prospect.txt", sb.ToString());
                sb.Clear();
                MessageBox.Show(e.Message, "Error",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Information);
            }
        }
        public static void SendProspectCron()//object source, ElapsedEventArgs e)
        {/*
            DateTime now1 = DateTime.Now;
            DateTime NowNoSecond2 = new DateTime(now1.Year, now1.Month, now1.Day, now1.Hour, now1.Minute, 0, 0);
            DateTime firstRun2 = new DateTime(NowNoSecond2.Year, NowNoSecond2.Month, NowNoSecond2.Day, 0, 0, 0, 0);
            if (NowNoSecond2 == firstRun2)
            {
                UtilsConfig.CRONSYNCHROPROSPECTDONE = "FALSE";
            }
            string[] date = UtilsConfig.CRONSYNCHROPROSPECT.ToString().Split(':');
            int hours, minutes;
            hours = Int32.Parse(date[0]);
            minutes = Int32.Parse(date[1]);
            DateTime now = DateTime.Now;
            DateTime NowNoSecond = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, 0);
            DateTime firstRun = new DateTime(now.Year, now.Month, now.Day, hours, minutes, 0, 0);
            if (NowNoSecond == firstRun && UtilsConfig.CRONSYNCHROPROSPECTDONE.ToString().Equals("FALSE"))
            {*/
                try
                {
                
                    //ZohoController.CreateProspect();
                    var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                    var clientsSageObj = compta.FactoryProspect.List;
                    var clients = GetListOfprospectToProcess(clientsSageObj);
                    List<string> prospects = ZohoController.GetListOfprospectToProcessZoho();
                    //int increm = 100 / clients.Count;
                    File.AppendAllText("Log\\SendProspect.txt", "will create  : " + clients.Count.ToString() + " Prospect" + Environment.NewLine);
                    foreach (Prospect client in clients)
                    {
                        if (!prospects.Contains(client.CT_NUM))
                        {
                            if (!String.IsNullOrEmpty(client.Email))
                            {
                                if (!IsValidEmail(client.Email))
                                {
                                    File.AppendAllText("Log\\SendProspect.txt", "adresse mail invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    client.Email = "";
                                }
                            }
                            if (!String.IsNullOrEmpty(client.telephone))
                            {
                                if (!validTelephoneNo(client.telephone))
                                {
                                    File.AppendAllText("Log\\SendProspect.txt", "Telephone invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                client.telephone = "";
                                }
                            }
                            if (!String.IsNullOrEmpty(client.telecopie))
                            {
                                if (!validTelephoneNo(client.telecopie))
                                {
                                    File.AppendAllText("Log\\SendProspect.txt", "Telecopie invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                client.telecopie = "";
                                    //continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.Website))
                            {
                                if (IsValidEmail(client.Website))
                                {
                                    File.AppendAllText("Log\\SendProspect.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                client.Website = "";
                                //continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.Website))
                            {
                                if (validTelephoneNo(client.Website))
                                {
                                    File.AppendAllText("Log\\SendProspect.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                client.Website = "";
                            }
                            }
                            //string articleXML = UtilsSerialize.SerializeObject(client);
                            File.AppendAllText("Log\\SendProspect.txt", "try to send : " + client.Intitule.ToString() + Environment.NewLine);
                            ZohoController.SendProspect(client);
                            File.AppendAllText("Log\\SendProspect.txt", "Done : " + client.CT_NUM.ToString() + Environment.NewLine);
                        }
                        else
                        {
                            ZohoController.UpdateProspect(client);
                        }
                    }
                    List<ZCRMRecord> leads = ZohoController.GetRecordsOfprospectToProcessZoho();
                    foreach (ZCRMRecord lead in leads)
                    {
                        ZohoController.CreateProspect(lead);
                    }
                SendClientCron();
                }
                catch (Exception s)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\CronProspectErreur.txt", sb.ToString());
                    sb.Clear();

                }
            //}
        }
        public static void SendClientCron()//object source, ElapsedEventArgs e)
        {
            
                try
                {
                    //ZohoController.CreateProspect();
                    /*var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                    var clientsSageObj = compta.FactoryClient.List;
                    var clients = GetListOfClientToProcess(clientsSageObj);
                    List<string> Accounts = ZohoController.GetListOfClientToProcessZoho();
                    //int increm = 100 / clients.Count;
                    File.AppendAllText("Log\\SendClient.txt", "will create  : " + clients.Count.ToString() + " Clients" + Environment.NewLine);
                    /*foreach (Client client in clients)
                    {
                        if (!Accounts.Contains(client.CT_NUM))
                        {
                            if (!String.IsNullOrEmpty(client.Email))
                            {
                                if (!IsValidEmail(client.Email))
                                {
                                    File.AppendAllText("Log\\SendClient.txt", "adresse mail invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.telephone))
                            {
                                if (!validTelephoneNo(client.telephone))
                                {
                                    File.AppendAllText("Log\\SendClient.txt", "Telephone invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.telecopie))
                            {
                                if (!validTelephoneNo(client.telecopie))
                                {
                                    File.AppendAllText("Log\\SendClient.txt", "Telecopie invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.Website))
                            {
                                if (IsValidEmail(client.Website))
                                {
                                    File.AppendAllText("Log\\SendClient.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    continue;
                                }
                            }
                            if (!String.IsNullOrEmpty(client.Website))
                            {
                                if (validTelephoneNo(client.Website))
                                {
                                    File.AppendAllText("Log\\SendClient.txt", "site invalide : " + client.CT_NUM.ToString() + Environment.NewLine);
                                    continue;
                                }
                            }
                            //string articleXML = UtilsSerialize.SerializeObject(client);
                            File.AppendAllText("Log\\SendClient.txt", "try to send : " + client.Intitule.ToString() + Environment.NewLine);
                            ZohoController.SendClient(client);
                            File.AppendAllText("Log\\SendClient.txt", "Done : " + client.CT_NUM.ToString() + Environment.NewLine);
                        }
                        else
                        {
                            ZohoController.UpdateClient(client);
                        }
                    }*/
                    List<ZCRMRecord> leads = ZohoController.GetRecordsOfAccountsToProcessZoho();
                    foreach (ZCRMRecord lead in leads)
                    {
                        ZohoController.TransformToCLient(lead.Data["CT_Num"].ToString(), lead.Data["id"].ToString());
                    }
                }
                catch (Exception s)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                    sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                    File.AppendAllText("Log\\TransformErreur.txt", sb.ToString());
                    sb.Clear();

                }
            }
        
        private static List<Client> GetClientToProcess(IBOClient3 clientsSageObj)
        {
            List<Client> clientToProcess = new List<Client>();

            Client client = new Client(clientsSageObj);

            if (!HandleClientError(client))
            {
                client.setClientLivraisonAdresse();
                clientToProcess.Add(client);
            }
           
            return clientToProcess;
        }
        private static List<Prospect> GetProspectToProcess(IBOClient3 clientsSageObj)
        {
            List<Prospect> clientToProcess = new List<Prospect>();

            Prospect client = new Prospect(clientsSageObj);

            if (!HandleProspectError(client))
            {
                client.setClientLivraisonAdresse();
                clientToProcess.Add(client);
            }

            return clientToProcess;
        }
        private static List<Prospect> GetListOfprospectToProcess(IBICollection clientsSageObj)
        {
            List<Prospect> clientToProcess = new List<Prospect>();

                foreach (IBOClient3 clientSageObj in clientsSageObj)
                {
                Prospect client = new Prospect(clientSageObj);

                    if (!HandleProspectError(client))
                    {
                        client.setClientLivraisonAdresse();
                        clientToProcess.Add(client);
                    }
                }
            return clientToProcess;
        }
        /// <summary>
        /// Permet de vérifier si un client comporte des erreur ou non
        /// </summary>
        /// <param name="client">Client à tester</param>
        /// <returns></returns>
        private static bool HandleClientError(Client client)
        {
            bool error = false;
            /*
            if(String.IsNullOrEmpty(client.Email))
            {
                error = true;
                return error;
               // SingletonUI.Instance.LogBox.Invoke((MethodInvoker)(() => SingletonUI.Instance.LogBox.AppendText("Client :  " + client.Intitule + " No Mail Found" + Environment.NewLine)));


                // on affiche une erreur + log 
            }
            */

            return error;
        }
        private static bool HandleProspectError(Prospect client)
        {
            bool error = false;
            /*
            if (String.IsNullOrEmpty(client.Email))
            {
                error = true;
                return error;
                // SingletonUI.Instance.LogBox.Invoke((MethodInvoker)(() => SingletonUI.Instance.LogBox.AppendText("Client :  " + client.Intitule + " No Mail Found" + Environment.NewLine)));


                // on affiche une erreur + log 
            }

            */
            return error;
        }

        /// <summary>
        /// Permet de récupérer une liste de Client depuis une liste de Client SAGE
        /// </summary>
        /// <param name="clientsSageObj">List de client SAGE</param>
        /// <returns></returns>
        private static List<Client> GetListOfClientToProcess(IBICollection clientsSageObj)
        {
            File.AppendAllText("Log\\AllClientTest.txt", "Worked!!!!3" + Environment.NewLine);
            List<Client> clientToProcess = new List<Client>();
                foreach (IBOClient3 clientSageObj in clientsSageObj)
                {
                    Client client = new Client(clientSageObj);
                //File.AppendAllText("Log\\AllClientTest.txt", "Worked CTNUM : " + client.CT_NUM + Environment.NewLine);
                //File.AppendAllText("Log\\AllClientTest.txt", "client count : "+ clientToProcess.Count.ToString() + Environment.NewLine);
                if (!HandleClientError(client))
                    {
                        client.setClientLivraisonAdresse();
                        clientToProcess.Add(client);
                    }
                }
            return clientToProcess;
        }

        /// <summary>
        /// Permet de vérifier si un Client existe dans SAGE
        /// </summary>
        /// <param name="CT_num"></param>
        /// <returns></returns>
        public static bool CheckIfClientExist(string CT_num)
        {
            if (String.IsNullOrEmpty(CT_num))
            {
                return false;
            }
            else
            {
                var compta = SingletonConnection.Instance.Gescom.CptaApplication;
                if (compta.FactoryClient.ExistNumero(CT_num))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }

        }

        /// <summary>
        /// Permet de crée un Client dans la base SAGE depuis un objet json de prestashop
        /// </summary>
        /// <param name="jsonClient">json du Client à crée</param>
        /// <returns></returns>
        public static string CreateNewClient(string jsonClient, JToken jsonOrder)
        {
            JObject customer = JObject.Parse(jsonClient);

            var compta = SingletonConnection.Instance.Gescom.CptaApplication;
            var gescom = SingletonConnection.Instance.Gescom;
            IBOClient3 clientSage = (IBOClient3)compta.FactoryClient.Create();
            clientSage.SetDefault();

            if (jsonOrder["invoice_adresse1"].ToString().Length > 35)
            {
                clientSage.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString().Substring(0, 35);
            }
            else
            {
                clientSage.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString();
            }
            clientSage.Adresse.Complement = jsonOrder["invoice_adresse2"].ToString();
            clientSage.Adresse.CodePostal = jsonOrder["invoice_postcode"].ToString();
            clientSage.Adresse.Ville = jsonOrder["invoice_city"].ToString();
            clientSage.Adresse.Pays = jsonOrder["invoice_country"].ToString();
            clientSage.Telecom.Telephone = jsonOrder["invoice_phone"].ToString();

            if (String.IsNullOrEmpty(UtilsConfig.PrefixClient))
            {
                // pas de configuration renseigner pour le prefix client
                // todo log
                int iterID = Int32.Parse(UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Client.Value, "getClientIterationSage&clientID="+ customer["id"].ToString()));
                while (compta.FactoryClient.ExistNumero(iterID.ToString()))
                {
                    iterID++;
                }
                clientSage.CT_Num = iterID.ToString();
            }
            else
            {
                clientSage.CT_Num = UtilsConfig.PrefixClient + customer["id"].ToString();
            }
            if (String.IsNullOrEmpty(UtilsConfig.CatTarif))
            {
                // pas de configuration renseigner pour la cat tarif par defaut
                // todo log
            }
            else
            {
                clientSage.CatTarif = gescom.FactoryCategorieTarif.ReadIntitule(UtilsConfig.CatTarif);
            }
            if (String.IsNullOrEmpty(UtilsConfig.CompteG))
            {
                // pas de configuration renseigner pour la cat tarif par defaut
                // todo log
            }
            else
            {
                clientSage.CompteGPrinc = compta.FactoryCompteG.ReadNumero(UtilsConfig.CompteGnum);
            }

            string intitule = customer["firstname"].ToString().ToUpper() + " " + customer["lastname"].ToString().ToUpper();

            clientSage.Telecom.EMail = customer["email"].ToString();
            clientSage.CT_Intitule = intitule;
            if (intitule.Length > 17)
            {
                clientSage.CT_Classement = intitule.Substring(0, 17);
            }
            else
            {
                clientSage.CT_Classement = intitule;
            }

            clientSage.Write();

            IBOClientLivraison3 addrprinc = (IBOClientLivraison3)clientSage.FactoryClientLivraison.Create();

            addrprinc.LI_Intitule = jsonOrder["invoice_name"].ToString();
            if (jsonOrder["invoice_adresse1"].ToString().Length > 35)
            {
                addrprinc.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString().Substring(0, 35);
            }
            else
            {
                addrprinc.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString();
            }
            addrprinc.Adresse.Complement = jsonOrder["invoice_adresse2"].ToString();
            addrprinc.Adresse.CodePostal = jsonOrder["invoice_postcode"].ToString();
            addrprinc.Adresse.Ville = jsonOrder["invoice_city"].ToString();
            addrprinc.Adresse.Pays = jsonOrder["invoice_country"].ToString();
            if (String.IsNullOrEmpty(UtilsConfig.CondLivraison))
            {
                // pas de configuration renseigner pour CondLivraison par defaut
                // todo log
            }
            else
            {
                addrprinc.ConditionLivraison = gescom.FactoryConditionLivraison.ReadIntitule(UtilsConfig.CondLivraison);
            }
            if (String.IsNullOrEmpty(UtilsConfig.Expedition))
            {
                // pas de configuration renseigner pour Expedition par defaut
                // todo log
            }
            else
            {
                addrprinc.Expedition = gescom.FactoryExpedition.ReadIntitule(UtilsConfig.Expedition);
            }
            clientSage.LivraisonPrincipal = addrprinc;
            addrprinc.Write();

            // on envoie une notification à préstashop pour lui informer de la créeation dans SAGE du client
            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Client.Value, "updateCTnum&clientID=" + customer["id"].ToString() + "&ct_num=" + clientSage.CT_Num);
            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Client.Value, "updateIter&iter=" + clientSage.CT_Num);


            return clientSage.CT_Num;
        }

    }
}
