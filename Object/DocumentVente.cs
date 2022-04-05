using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebservicesSage.Object
{
    public class DocumentVente
    {
        public string NumPiece { get; set; }
        public string Devise { get; set; }
        public string Reference { get; set; }
        public List<LigneDocumentVente> DocLigne { get; set; }
        public string Client { get; set; }
        public DateTime DateCreation { get; set; }
        public string DocumentType { get; set; }
        public double Total_HT { get; set; }

        public DocumentVente(string numPiece, string devise, string reference,string client,string documentType,DateTime dateCreation, double total_HT)
        {
            NumPiece = numPiece;
            Devise = devise;
            Reference = reference;
            Client = client;
            DateCreation = dateCreation;
            DocumentType = documentType;
            DocLigne = null;
            Total_HT = total_HT;
        }
        public DocumentVente()
        { }
    }
    
}
