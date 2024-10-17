using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Letter_Maker
{
    public interface IStrategy
    {
        bool Check(string selectedPath);
    }
    public class ADKStrategy : IStrategy
    {
        public bool Check(string selectedPath)
        {
            return true;
        }
    }
}
