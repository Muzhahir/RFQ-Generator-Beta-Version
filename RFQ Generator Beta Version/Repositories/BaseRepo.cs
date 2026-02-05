using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System.Repositories
{
    public class BaseRepo<T> where T : class, new()
    {
        protected string connectionString =
    @"Data Source=(LocalDB)\MSSQLLocalDB;
      Initial Catalog=RFQDB;
      Integrated Security=True";

    }
}