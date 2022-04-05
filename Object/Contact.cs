using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Objets100cLib;

namespace WebservicesSage.Object
{
    
    public class Contact
    {
        public string Civilite { get; set; }
        public string Nom { get; set; }
        public string Prenom { get; set; }
        public string Service { get; set; }
        public string Fonction { get; set; }
        public string Telephone { get; set; }
        public string Portable { get; set; }
        public string Telecopie { get; set; }
        public string Skype { get; set; }
        public string LinkedIn { get; set; }
        public string Email { get; set; }
        public string Facebook { get; set; }
        public Contact(IBOTiersContact3 contact)
        {
            Civilite = contact.Civilite.ToString();
            Nom = contact.Nom;
            Prenom = contact.Prenom;
            Service = contact.ServiceContact.S_Intitule;
            Fonction = contact.Fonction;
            Telephone = contact.Telecom.Telephone;
            Portable = contact.Telecom.Portable;
            Telecopie = contact.Telecom.Telecopie;
            Skype = contact.Skype;
            LinkedIn = contact.LinkedIn;
            Email = contact.Telecom.EMail;
            Facebook = contact.Facebook;
        }
    }
}
