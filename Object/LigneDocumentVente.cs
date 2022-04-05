using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebservicesSage.Object
{
    public class LigneDocumentVente
    {
        public string RefArticle { get; set; }
        public string Designiation { get; set; }
        public string PriceHT { get; set; }
        public string Quantity { get; set; }
        public string Devise { get; set; }

        public LigneDocumentVente(string RefArticle,string Designiation, string priceHT,string quantity,string devise)
        {
            this.RefArticle = RefArticle;
            this.Designiation = Designiation;
            this.PriceHT = priceHT;
            this.Quantity = quantity;
            this.Devise = devise;
        }
        public LigneDocumentVente()
        {

        }
    }

}
