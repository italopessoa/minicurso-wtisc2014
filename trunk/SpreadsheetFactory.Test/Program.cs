using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetFactory.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Header header = new Header();
            header.AddFilter("nome", "italo");
            header.AddFilter("idade", 1);
            header.AddFilter("nome", "jose");
        }
    }
}
