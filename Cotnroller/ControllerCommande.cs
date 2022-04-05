using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using WebservicesSage.Utils;
using WebservicesSage.Utils.Enums;
using WebservicesSage.Singleton;
using Objets100cLib;
using WebservicesSage.Object;
using WebservicesSage.Object.DBObject;
using System.Windows.Forms;
using LiteDB;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;

namespace WebservicesSage.Cotnroller
{
    public static class ControllerCommande
    {

        /// <summary>
        /// Lance le service de check des nouvelles commandes prestashop
        /// Définir le temps de passage de la tâche dans la config
        /// </summary>
        public static void LaunchService()
        {
            // SingletonUI.Instance.LogBox.Invoke((MethodInvoker)(() => SingletonUI.Instance.LogBox.AppendText("Commande Services Launched " + Environment.NewLine)));
           /* System.Timers.Timer timer = new System.Timers.Timer();
            timer.Elapsed += new ElapsedEventHandler(CheckForNewOrder);
            timer.Interval = UtilsConfig.CronTaskCheckForNewOrder;
            timer.Enabled = true;
            
            System.Timers.Timer timerUpdateStatut = new System.Timers.Timer();
            timerUpdateStatut.Elapsed += new ElapsedEventHandler(UpdateStatuOrder);
            timerUpdateStatut.Interval = UtilsConfig.CronTaskUpdateStatut;
            timerUpdateStatut.Enabled = true;*/
        }

        /// <summary>
        /// Event levé par une nouvelle commande dans prestashop
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        public static void CheckForNewOrder(object source, ElapsedEventArgs e)
        {

            string currentIdOrder = "0";
            try
            {
                string response = UtilsWebservices.SendData(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "checkOrder");

                if (!response.Equals("none") && !response.Equals("[]"))
                {
                    JArray orders = JArray.Parse(response);
                    List<int> currentCustomer_IDs = new List<int>();
                    foreach (var order in orders)
                    {
                        currentIdOrder = order["id_order"].ToString();
                        if (ControllerClient.CheckIfClientExist(order["ALTAIS_CT_NUM"].ToString()))
                        {
                            // si le client existe on associé la commande à son compte
                            try
                            {
                                AddNewOrderForCustomer(order, order["ALTAIS_CT_NUM"].ToString());
                            }
                            catch (Exception exception)
                            {
                                File.AppendAllText("Log\\exception.txt", exception.ToString());

                            }
                            
                        }
                        
                       else
                        {
                            // si le client n'existe pas on récupère les info de presta et on le crée dans la base sage 
                            string client = UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Client.Value, "getClient&clientID=" + order["id_customer"]);
                            string ct_num = ControllerClient.CreateNewClient(client, order);

                            if (!String.IsNullOrEmpty(ct_num))
                            {
                                // le client à bien été crée on peut intégrer la commande sur son compte sage
                                AddNewOrderForCustomer(order, ct_num);
                            }
                        }
                    }
                }
            }
            catch (Exception s)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now+ s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                UtilsMail.SendErrorMail(DateTime.Now + s.Message + Environment.NewLine + s.StackTrace + Environment.NewLine, "COMMANDE");
                File.AppendAllText("Log\\order.txt", sb.ToString());
                sb.Clear();
                UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "errorOrder&orderID=" + currentIdOrder);
            }

        }

        public static void UpdateStatuOrder(object source, ElapsedEventArgs e)
        {
            try
            {
                File.AppendAllText("Log\\time.txt", "start :" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                var gescom = SingletonConnection.Instance.Gescom;
                var compta = SingletonConnection.Instance.Compta;

                /***********/
                //test
                //IBICollection AllOrderstest = gescom.FactoryDocumentVente.;

                /***********/
                IBICollection AllOrders = gescom.FactoryDocumentVente.List;
                using (var db = new LiteDatabase(@"MyData.db"))
                {
                    string prestaStatutId, orderStatutId, statut1, statut2, statut3;
                    string[] prestaID, orderStatut, docstatut1, docstatut2, docstatut3;
                    UtilsConfig.PrestaStatutId.TryGetValue("default", out prestaStatutId);
                    prestaID = prestaStatutId.Split('_');
                    UtilsConfig.OrderMapping.TryGetValue("default", out orderStatutId);
                    orderStatut = orderStatutId.Split('_');
                    UtilsConfig.OrderMapping.TryGetValue(orderStatut[0], out statut1);
                    UtilsConfig.OrderMapping.TryGetValue(orderStatut[1], out statut2);
                    UtilsConfig.OrderMapping.TryGetValue(orderStatut[2], out statut3);
                    docstatut1 = statut1.Split('_');
                    docstatut2 = statut2.Split('_');
                    docstatut3 = statut3.Split('_');
                    // Get a collection (or create, if doesn't exist)
                    var col = db.GetCollection<LinkedCommandeDB>("Commande");
                    foreach (LinkedCommandeDB item in col.FindAll())
                    {
                        string SQL = null;
                        SQL = "SELECT  DO_Piece FROM ["+ ConfigurationManager.AppSettings["DBNAME"].ToString()+"].[dbo].[F_DOCENTETE] WHERE DO_Ref = '" +item.DO_Ref+"' And DO_Domaine = 0";
                        SqlDataReader dataReader = DB.Select(SQL);
                        while (dataReader.Read())
                        {
                            string data = dataReader.GetValue(0).ToString();
                            if ((data.Substring(0, 2)).Equals("PL"))
                            {
                                IBODocumentVente3 order = gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVentePrepaLivraison, data);
                                update(order, item, col, docstatut1, docstatut2, docstatut3, prestaID);
                            }else if ((data.Substring(0,2)).Equals("BL"))
                            {
                                IBODocumentVente3 order = gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteLivraison, data);
                                update(order, item, col, docstatut1, docstatut2, docstatut3, prestaID);
                            }
                            else if ((data.Substring(0, 2)).Equals("FA"))
                            {
                                IBODocumentVente3 order = gescom.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFacture, data);
                                update(order, item, col, docstatut1, docstatut2, docstatut3, prestaID);
                            }

                            /*}
                             /*   foreach (IBODocumentVente3 order in AllOrders)
                            {
                            if (order.DO_Ref.Equals(item.DO_Ref))
                            {
                                File.AppendAllText("Log\\time.txt", "find : "+ order.DO_Ref.ToString() + " "+ DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                File.AppendAllText("Log\\time.txt", "order.do_type : " + order.DO_Type.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                File.AppendAllText("Log\\time.txt", "item.do_type : " + item.OrderType.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                File.AppendAllText("Log\\time.txt", "staut1 : " + docstatut1[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                File.AppendAllText("Log\\time.txt", "staut2 : " + docstatut2[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                File.AppendAllText("Log\\time.txt", "staut3 : " + docstatut3[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);

                                if (!(order.DO_Type.ToString().Equals(item.OrderType)))
                                {
                                    if (order.DO_Type.ToString().Equals("DocumentTypeVenteCommande"))
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        File.AppendAllText("Log\\time.txt", "DO_type : " + order.DO_Type.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                                        if (order.DO_Type.ToString().Equals(docstatut1[0]))
                                        {
                                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[0]);
                                            item.OrderType = docstatut1[0];
                                            col.Update(item);

                                            break;
                                        }
                                        if (order.DO_Type.ToString().Equals(docstatut2[0]))
                                        {
                                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[1]);
                                            item.OrderType = docstatut2[0];
                                            col.Update(item);
                                            break;
                                        }
                                        if (order.DO_Type.ToString().Equals(docstatut3[0]))
                                        {
                                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[2]);
                                            col.Delete(item.Id);
                                            break;
                                        }
                                    }
                                }

                            }*/
                        }
                    }
                }
                File.AppendAllText("Log\\time.txt", "end :"+DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
            }
            catch (Exception ex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + ex.Message + Environment.NewLine);
                sb.Append(DateTime.Now + ex.StackTrace + Environment.NewLine);
                File.AppendAllText("Log\\statut.txt", sb.ToString());
                sb.Clear();
            }
            
        }

        private static void update(IBODocumentVente3 order,LinkedCommandeDB item, LiteCollection<LinkedCommandeDB> col , string[] docstatut1, string[] docstatut2, string[] docstatut3, string[] prestaID)
        {
            if (order.DO_Ref.Equals(item.DO_Ref))
            {
                File.AppendAllText("Log\\time.txt", "find : " + order.DO_Ref.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                File.AppendAllText("Log\\time.txt", "order.do_type : " + order.DO_Type.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                File.AppendAllText("Log\\time.txt", "item.do_type : " + item.OrderType.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                File.AppendAllText("Log\\time.txt", "staut1 : " + docstatut1[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                File.AppendAllText("Log\\time.txt", "staut2 : " + docstatut2[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                File.AppendAllText("Log\\time.txt", "staut3 : " + docstatut3[0].ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);

                if (!(order.DO_Type.ToString().Equals(item.OrderType)))
                {
                    if (order.DO_Type.ToString().Equals("DocumentTypeVenteCommande"))
                    {
                        //break;
                    }
                    else
                    {
                        File.AppendAllText("Log\\time.txt", "DO_type : " + order.DO_Type.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine);
                        if (order.DO_Type.ToString().Equals(docstatut1[0]))
                        {
                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[0]);
                            item.OrderType = docstatut1[0];
                            col.Update(item);

                          //  break;
                        }
                        if (order.DO_Type.ToString().Equals(docstatut2[0]))
                        {
                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[1]);
                            item.OrderType = docstatut2[0];
                            col.Update(item);
                            //break;
                        }
                        if (order.DO_Type.ToString().Equals(docstatut3[0]))
                        {
                            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + item.OrderID + "&orderType=" + prestaID[2]);
                            col.Delete(item.Id);
                            //break;
                        }
                    }
                }

            }

        }

        /// <summary>
        /// Crée une nouvelle commande pour un utilisateur
        /// </summary>
        /// <param name="jsonOrder">Order à crée</param>
        /// <param name="CT_Num">Client</param>
        public static void AddNewOrderForCustomer(JToken jsonOrder, string CT_Num)
        {
            var gescom = SingletonConnection.Instance.Gescom;
            
            // création de l'entête de la commande 

            IBOClient3 customer = gescom.CptaApplication.FactoryClient.ReadNumero(CT_Num);
            Client client = new Client(customer);
            IBODocumentVente3 order = gescom.FactoryDocumentVente.CreateType(DocumentType.DocumentTypeVenteCommande);
            order.SetDefault();
            order.SetDefaultClient(customer);
            order.DO_Date = DateTime.Now;
            try
            {
                string carrier_id = jsonOrder["order_carriere"].ToString();
                order.Expedition = gescom.FactoryExpedition.ReadIntitule(UtilsConfig.OrderCarrierMapping[carrier_id]);
            }
            catch (Exception s){
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + s.Message + Environment.NewLine);
                sb.Append(DateTime.Now + s.StackTrace + Environment.NewLine);
                UtilsMail.SendErrorMail(DateTime.Now + s.Message + Environment.NewLine + s.StackTrace + Environment.NewLine, "TRANSPORTEUR");
                File.AppendAllText("Log\\order.txt", sb.ToString());
                sb.Clear();
            }
            order.Souche = gescom.FactorySoucheVente.ReadIntitule(UtilsConfig.Souche);
            order.DO_Ref =jsonOrder["reference"].ToString();
            order.SetDefaultDO_Piece();

            // on définis l'adresse de livraison de la commande

            bool asAdressMatch = false;
            string intitule = "* " + jsonOrder["shipping_company"].ToString() + " " + jsonOrder["shipping_postcode"].ToString();
            foreach (IBOClientLivraison3 adress in customer.FactoryClientLivraison.List)
            {
                if (intitule.Length >35)
                {
                     intitule = intitule.Substring(0, 35);
                }
                if (adress.LI_Intitule.Equals(intitule))
                {
                    if (jsonOrder["shipping_adresse1"].ToString().Length > 35)
                    {
                        adress.Adresse.Adresse = jsonOrder["shipping_adresse1"].ToString().Substring(0, 35);
                    }
                    else
                    {
                        adress.Adresse.Adresse = jsonOrder["shipping_adresse1"].ToString();
                    }
                    string contact = jsonOrder["shipping_firstname"].ToString() + " " + jsonOrder["shipping_lastname"].ToString();
                    if (contact.Length > 35)
                    {
                        adress.LI_Contact = contact.Substring(0, 35);
                    }
                    else
                    {
                        adress.LI_Contact = contact;
                    }
                    adress.Adresse.Complement = jsonOrder["shipping_adresse2"].ToString();
                    adress.Adresse.CodePostal = jsonOrder["shipping_postcode"].ToString();
                    adress.Adresse.Ville = jsonOrder["shipping_city"].ToString();
                    adress.Adresse.Pays = jsonOrder["shipping_country"].ToString();
                    adress.Adresse.CodeRegion = "WEB";
                    adress.Telecom.EMail = jsonOrder["addr_email"].ToString();
                    adress.Telecom.Telephone = jsonOrder["shipping_phone"].ToString();
                    adress.Write();
                    order.LieuLivraison = adress;
                    asAdressMatch = true;
                    break;
                }
            }

            if (!asAdressMatch)
            {
                // si on a trouver aucune adresse coresspondant e sur le client alors on la crée
                IBOClientLivraison3 adress = (IBOClientLivraison3)customer.FactoryClientLivraison.Create();
                adress.SetDefault();
                if (intitule.Length > 35)
                {
                    intitule = intitule.Substring(0, 35);
                }
                adress.LI_Intitule = intitule.ToString();
                if (jsonOrder["shipping_adresse1"].ToString().Length > 35)
                {
                    adress.Adresse.Adresse = jsonOrder["shipping_adresse1"].ToString().Substring(0, 35);
                }
                else
                {
                    adress.Adresse.Adresse = jsonOrder["shipping_adresse1"].ToString();
                }
                string contact = jsonOrder["shipping_firstname"].ToString() + " " + jsonOrder["shipping_lastname"].ToString();
                if (contact.Length>35)
                {
                    adress.LI_Contact = contact.Substring(0, 35);
                }
                else
                {
                    adress.LI_Contact = contact;
                }
                adress.Adresse.Complement = jsonOrder["shipping_adresse2"].ToString();
                adress.Adresse.CodePostal = jsonOrder["shipping_postcode"].ToString();
                adress.Adresse.Ville = jsonOrder["shipping_city"].ToString();
                adress.Adresse.Pays = jsonOrder["shipping_country"].ToString();
                adress.Adresse.CodeRegion = "WEB";
                adress.Telecom.EMail = jsonOrder["addr_email"].ToString();
                adress.Telecom.Telephone = jsonOrder["shipping_phone"].ToString();
                if (String.IsNullOrEmpty(UtilsConfig.CondLivraison))
                {
                    // pas de configuration renseigner pour CondLivraison par defaut
                    // todo log
                }
                else
                {
                    adress.ConditionLivraison = gescom.FactoryConditionLivraison.ReadIntitule(UtilsConfig.CondLivraison);
                }
                if (String.IsNullOrEmpty(UtilsConfig.Expedition))
                {
                    // pas de configuration renseigner pour Expedition par defaut
                    // todo log
                }
                else
                {
                    adress.Expedition = gescom.FactoryExpedition.ReadIntitule(UtilsConfig.Expedition);
                }
                adress.Write();

                 // on ajoute une adresse par defaut sur la fiche client si il y en a pas
                 /*
                 if (String.IsNullOrEmpty(customer.Adresse.Adresse))
                 {
                     if (jsonOrder["invoice_adresse1"].ToString().Length > 35)
                     {
                         customer.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString().Substring(0, 35);
                     }
                     else
                     {
                         customer.Adresse.Adresse = jsonOrder["invoice_adresse1"].ToString();
                     }
                     customer.Adresse.Complement = jsonOrder["invoice_adresse2"].ToString();
                     customer.Adresse.CodePostal = jsonOrder["invoice_postcode"].ToString();
                     customer.Adresse.Ville = jsonOrder["invoice_city"].ToString();
                     customer.Adresse.Pays = jsonOrder["invoice_country"].ToString();
                     customer.Telecom.Telephone = jsonOrder["invoice_phone"].ToString();
                     customer.Write();
                 }*/
                 
                    order.LieuLivraison = adress;
             }


            //order.LieuLivraison.Adresse.Adresse = "test";
            //order.LieuLivraison.Client. = customer;
            //order.SetDefaultClient(customer);
            //order.Write();

            order.Write();

            //order.InfoLibre[1] = "test";
            // création des lignes de la commandes
            try
            {
                Boolean dispo = false;
                foreach (JToken product in jsonOrder["products"].Children())
                {
                    // on récupère la chaine de gammages d'un produit
                    string product_attribut_string = product["product_attribute_id_string"].ToString();
                    String[] subgamme = product_attribut_string.Split('|');

                    if (product["product_ref"].ToString().Equals("noInsert"))
                    {
                        if (product["id_carrier"].ToString().Equals(UtilsConfig.IDTRANSPORT.ToString()))
                        {
                            dispo = true;
                        }
                        continue;
                    }
                    IBODocumentLigne3 docLigne = (IBODocumentLigne3)order.FactoryDocumentLigne.Create();
                    IBOArticle3 article = gescom.FactoryArticle.ReadReference(product["product_ref"].ToString());
                    docLigne.DL_PrixUnitaire = double.Parse(product["price"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                    if (subgamme.Length == 2)
                    {
                       
                        // produit à simple gamme
                        IBOArticleGammeEnum3 articleEnum = ControllerArticle.GetArticleGammeEnum1(article, new Gamme(subgamme[0], subgamme[1]));
                        docLigne.SetDefaultArticleMonoGamme(articleEnum, Int32.Parse(product["product_quantity"].ToString()));
                    }
                    else if (subgamme.Length == 4)
                    {
                        // produit à double gamme
                        IBOArticleGammeEnum3 articleEnum = ControllerArticle.GetArticleGammeEnum1(article, new Gamme(subgamme[0], subgamme[1], subgamme[2], subgamme[3]));
                        IBOArticleGammeEnum3 articleEnum2 = ControllerArticle.GetArticleGammeEnum2(article, new Gamme(subgamme[0], subgamme[1], subgamme[2], subgamme[3]));
                        docLigne.SetDefaultArticleDoubleGamme(articleEnum, articleEnum2, Int32.Parse(product["product_quantity"].ToString()));
                    }
                    else
                    {
                        //TODO mapping transporteur else use default transporter
                        if (product["product_ref"].ToString().Equals("noInsert"))
                        {
                            if (product["id_carrier"].ToString().Equals(UtilsConfig.IDTRANSPORT.ToString()))
                            {
                                dispo = true;
                            }
                            continue;
                        }
                        // produit simple
                        docLigne.SetDefaultArticle(gescom.FactoryArticle.ReadReference(product["product_ref"].ToString()), Int32.Parse(product["product_quantity"].ToString()));
                        docLigne.DL_PrixUnitaire = Convert.ToDouble(product["price"].ToString().Replace('.', ','));
                        if (product["id_carrier"].ToString().Equals(UtilsConfig.IDTRANSPORT.ToString()))
                        {
                            docLigne.DL_PrixUnitaire = Convert.ToDouble(product["price"].ToString().Replace('.', ','));
                            dispo = true;

                        }
                        else if (product["product_ref"].ToString().Equals("REMISE"))
                        {
                            docLigne.DL_PrixUnitaire = Convert.ToDouble(product["price"].ToString().Replace('.', ','));
                        }
                        
                    }

                    docLigne.Write();
                }
                if (!String.IsNullOrEmpty(jsonOrder["message"].ToString()))
                {
                    IBODocumentLigne3 docLigne = (IBODocumentLigne3)order.FactoryDocumentLigne.Create();
                    docLigne.SetDefaultArticle(gescom.FactoryArticle.ReadReference(UtilsConfig.Comment), Int32.Parse("1"));
                    docLigne.TxtComplementaire = jsonOrder["message"].ToString();
                    docLigne.DL_PrixUnitaire = Convert.ToDouble("0");
                    docLigne.Write();
                }
                /********     Order infolibre     ********/
                if (dispo)
                {
                    order.InfoLibre[1] = "WEB / dispo";
                }
                else
                {
                    order.InfoLibre[1] = "WEB";
                }
                order.InfoLibre[2] = jsonOrder["other"].ToString();
                order.InfoLibre[3] = jsonOrder["etage"].ToString();
                order.Write();
                /********     Order infolibre     ********/
            }

            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(DateTime.Now + e.Message + Environment.NewLine);
                sb.Append(DateTime.Now + e.StackTrace + Environment.NewLine);
                UtilsMail.SendErrorMail(DateTime.Now + e.Message + Environment.NewLine + e.StackTrace + Environment.NewLine, "COMMANDE LIGNE");
                File.AppendAllText("Log\\order.txt", sb.ToString());
                sb.Clear();
                UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "errorOrder&orderID=" + jsonOrder["id_order"]);
                order.Remove();
                return;
            }
            addOrderToLocalDB(jsonOrder["id_order"].ToString(), order.Client.CT_Num, order.DO_Piece, order.DO_Ref);
            // on notfie prestashop que la commande à bien été crée dans SAGE
            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "validateOrder&orderID=" + jsonOrder["id_order"]);
        }

        private static void addOrderToLocalDB(string orderID, string CT_Num, string DO_piece, string DO_Ref)
        {
            // Open database (or create if doesn't exist)
            using (var db = new LiteDatabase(@"MyData.db"))
            {
                // Get a collection (or create, if doesn't exist)
                var col = db.GetCollection<LinkedCommandeDB>("Commande");

                // Create your new customer instance
                var commande = new LinkedCommandeDB
                {
                    OrderID = orderID,
                    OrderType = "DocumentTypeVenteCommande",
                    CT_Num = CT_Num,
                    DO_piece = DO_piece,
                    DO_Ref = DO_Ref

                };
                col.Insert(commande);
            }
        }

        public static string GetPrestaOrderStatutFromMapping(DocumentType orderSageType= DocumentType.DocumentTypeVenteCommande, string prestaStatutId="")
        {
            Array x = UtilsConfig.OrderMapping.ToArray();
            foreach (KeyValuePair<string, string> kvp in UtilsConfig.OrderMapping.ToArray())
            {
                string[] res = kvp.Value.ToString().Split('_');
                if (res[1].Equals(prestaStatutId))
                {
                    return res[0].ToString();
                }
            }
            return "";
        }

        public static void UpdateStatutOnPresta(string orderID, string newType)
        {
            UtilsWebservices.SendDataNoParse(UtilsConfig.BaseUrl + EnumEndPoint.Commande.Value, "updateOrder&orderID=" + orderID + "&orderType=" + newType);
        }
    }
}
