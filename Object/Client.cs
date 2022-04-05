using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Objets100cLib;
using WebservicesSage.Object;
using WebservicesSage.Cotnroller;

namespace WebservicesSage.Object
{
    [Serializable()]
    public class Client
    {
        public string Intitule { get; set; }
        public bool Sommeil { get; set; }
        public string Email { get; set; }
        public string GroupeTarifaireIntitule { get; set; }
        public string CT_NUM { get; set; }
        public string Interlocuteur { get; set; }
        public string CompteCollectif { get; set; }
        public string LinkedIn { get; set; }
        public string telecopie { get; set; }
        public string Website { get; set; }
        public string Facebook { get; set; }
        public string Siret { get; set; }
        public string CodeNAF { get; set; }
        public string IdTVA { get; set; }
        public string CentralAchat {get;set;}
        public string representant { get; set; }
        public string telephone { get; set; }
        public string qualite { get; set; }
        public string commentaire { get; set; }
        public DateTime ModificationTime { get; set; }
        public List<ClientLivraisonAdress> clientBillingAdresses { get; set; }
        public List<Contact> Contacts { get; set; }
        private IBOClient3 clientFC;
        public string ZohoEntityId { get; set; }
        public string groupe { get; set; }
        public string enseigne { get; set; }

        public Client(IBOClient3 clientFC)
        {
            clientBillingAdresses = new List<ClientLivraisonAdress>();
            Contacts = new List<Contact>();
            this.clientFC = clientFC;
            Email = clientFC.Telecom.EMail;
            Intitule = clientFC.CT_Intitule;
            Sommeil = clientFC.CT_Sommeil;
            ModificationTime = clientFC.DateModification;
            GroupeTarifaireIntitule = clientFC.CatTarif.CT_Intitule;
            CT_NUM = clientFC.CT_Num;
            Interlocuteur = clientFC.CT_Contact;
            CompteCollectif = clientFC.CompteGPrinc.CG_Intitule ;
            LinkedIn = clientFC.CT_LinkedIn;
            telecopie = clientFC.Telecom.Telecopie;
            Website = clientFC.Telecom.Site;
            Facebook = clientFC.CT_Facebook;
            Siret = clientFC.CT_Siret;
            CodeNAF = clientFC.CT_Ape;
            IdTVA = clientFC.CT_Identifiant;
            telephone = clientFC.Telecom.Telephone;
            qualite = clientFC.CT_Qualite;
            CentralAchat = "";
            if (clientFC.CentraleAchat != null)
            {
                CentralAchat = clientFC.CentraleAchat.CT_Intitule;
            }
            commentaire = "";
            if (!String.IsNullOrEmpty(clientFC.CT_Commentaire ))
            {
                commentaire = clientFC.CT_Commentaire;
            }
            representant = "";
            if (clientFC.Representant != null)
            {
                if (!String.IsNullOrEmpty(clientFC.Representant.Nom))
                {
                    representant = clientFC.Representant.Nom;
                }
                if (!String.IsNullOrEmpty(clientFC.Representant.Prenom))
                {
                    representant += clientFC.Representant.Prenom;
                }
            }

             
            foreach (IBOTiersContact3 item in clientFC.FactoryTiersContact.List)
            {
                Contact con = new Contact(item);
                Contacts.Add(con);
            }
            var infolibreField = Singleton.SingletonConnection.Instance.Compta.FactoryTiers.InfoLibreFields;
            int compteur = 1;

            groupe = "";
            enseigne = "";
            foreach (var infoLibreValue in clientFC.InfoLibre)
            {
                if (infolibreField[compteur].Name.Equals("ZohoEntityID"))
                {
                    ZohoEntityId = infoLibreValue.ToString();
                }
                if (infolibreField[compteur].Name.Equals("enseigne"))
                {
                    enseigne = infoLibreValue.ToString();
                }
                if (infolibreField[compteur].Name.Equals("groupe"))
                {
                    groupe = infoLibreValue.ToString();
                }
                compteur++;
            }
        }
        public Client()
        {

        }
        public void setClientLivraisonAdresse()
        {
            //clientLivraisonAdresses = ControllerClientLivraisonAdress.getAllClientLivraisonAdressToProcess(clientFC.FactoryClientLivraison.List);
            clientBillingAdresses = ControllerClientLivraisonAdress.GetClientAdress(clientFC);
        }
    }
}
