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
    public class Prospect
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

        public Prospect(IBOClient3 clientFC)
        {
            clientBillingAdresses = new List<ClientLivraisonAdress>();
            Contacts = new List<Contact>();
            this.clientFC = clientFC;
            Email = "";
            if (!String.IsNullOrEmpty(clientFC.Telecom.EMail))
            {
                Email = clientFC.Telecom.EMail;
            }
            
            Intitule = clientFC.CT_Intitule;
            Sommeil = clientFC.CT_Sommeil;
            ModificationTime = clientFC.DateModification;
            
            if (clientFC.CatTarif != null)
            {
                if (!String.IsNullOrEmpty(clientFC.CatTarif.CT_Intitule))
                {
                    GroupeTarifaireIntitule = clientFC.CatTarif.CT_Intitule;
                }
            }
            else
            {
                GroupeTarifaireIntitule = "";
            }
            CT_NUM = clientFC.CT_Num;
            Interlocuteur = clientFC.CT_Contact;
            /*if (clientFC.CompteGPrinc != null)
            {
                if (!String.IsNullOrEmpty(clientFC.CompteGPrinc.CG_Intitule ))
                {
                    CompteCollectif = clientFC.CompteGPrinc.CG_Intitule;
                }
                else
                {
                    CompteCollectif = "";
                }
            }
            else
            {
                CompteCollectif = "";
            }*/
            if (String.IsNullOrEmpty(clientFC.CT_LinkedIn))
            {
                LinkedIn = "";
            }
            else
            {
                LinkedIn = clientFC.CT_LinkedIn;
            }
            
            if (!String.IsNullOrEmpty(clientFC.Telecom.Telecopie))
            {
                telecopie = clientFC.Telecom.Telecopie;
            }
            else
            {
                telecopie = "";
            }
            
            
            if (String.IsNullOrEmpty(clientFC.Telecom.Site))
            {
                Website = "";
            }
            else
            {
                Website = clientFC.Telecom.Site;
            }
            if (String.IsNullOrEmpty(clientFC.CT_Facebook))
            {
                Facebook = "";
            }
            else
            {
                Facebook = clientFC.CT_Facebook;
            }
            if (String.IsNullOrEmpty(clientFC.CT_Siret))
            {
                Siret = "";
            }
            else
            {
                Siret = clientFC.CT_Siret;
            }
            if (String.IsNullOrEmpty(clientFC.CT_Ape))
            {
                CodeNAF = "";
            }
            else
            {
                CodeNAF = clientFC.CT_Ape;
            }
            
            
            if (!String.IsNullOrEmpty(clientFC.Telecom.Telephone))
            {
                telephone = clientFC.Telecom.Telephone;
            }
            else
            {
                telephone = "";
            }
            if (!String.IsNullOrEmpty(clientFC.CT_Identifiant))
            {
                IdTVA = clientFC.CT_Identifiant;
            }
            else
            {
                IdTVA = "";
            }
            if (!String.IsNullOrEmpty(clientFC.CT_Qualite))
            {
                qualite = clientFC.CT_Qualite;
            }
            else
            {
                qualite = "";
            }
            if (clientFC.CentraleAchat != null)
            {
                CentralAchat = clientFC.CentraleAchat.CT_Intitule;
            }
            else
            {
                CentralAchat = "";
            }
            commentaire = "";
            if (!String.IsNullOrEmpty(clientFC.CT_Commentaire))
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
        public Prospect()
        {

        }
        public void setClientLivraisonAdresse()
        {
            //clientLivraisonAdresses = ControllerClientLivraisonAdress.getAllClientLivraisonAdressToProcess(clientFC.FactoryClientLivraison.List);
            clientBillingAdresses = ControllerClientLivraisonAdress.GetClientAdress(clientFC);
        }
    }
}
