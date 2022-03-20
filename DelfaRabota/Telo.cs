using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DelfaRabota
{
    public class Telo
    {
        public string name { get; set; }
        public string route { get; set; }
        public string phone { get; set; }
        public string adress { get; set; }

        public Telo(string name, string route, string phone, string adress)
        {
            this.name = name;
            this.route = route;
            this.phone = phone;
            this.adress = adress;
        }
    }
}
