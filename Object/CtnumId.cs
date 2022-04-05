using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebservicesSage.Object
{
    public class CtnumId
    {
        public string ct_num { get; set; }
        public long id { get; set; }
        public CtnumId()
        {

        }
        public CtnumId(string ctnum, long id)
        {
            this.ct_num = ctnum;
            this.id = id;
        }
    }
}
