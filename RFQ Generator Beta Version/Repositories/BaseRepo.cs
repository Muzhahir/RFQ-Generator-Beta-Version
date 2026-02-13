using System.Configuration;

namespace RFQ_Generator_System.Repositories
{
    public abstract class BaseRepo<T> where T : class, new()
    {
        protected readonly string connectionString;

        protected BaseRepo()
        {
            connectionString =
                ConfigurationManager.ConnectionStrings["RFQ_DB"].ConnectionString;
        }
    }
}
