using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Letter_Maker
{
    public class Author
    {
        public string authorName { get; set; }
        public string phNumber { get; set; }
        internal Dictionary<string, string> spis = new Dictionary<string, string>();
        public Author() { }

        public void Add(string name,string phone)
        {
            spis.Add(name, phone);
        }
       
        public void Sort()
        {
            spis = spis.OrderBy(obj => obj.Key).ToDictionary(obj => obj.Key, obj => obj.Value);
        }
    }
}
