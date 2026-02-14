using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal class ConnectionDB
    {
        public string GetConnection()
        {
            string con = @"Data Source=LAPTOP-4AOSRHAE\\SQLEXPRESS;Integrated Security=True";
            return con;
        }
    }
}
