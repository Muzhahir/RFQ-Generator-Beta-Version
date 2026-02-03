using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System.Repositories
{
    public class ClientRepo : BaseRepo<Client>
    {
        // Get all clients
        public List<Client> GetAllClients()
        {
            var clients = new List<Client>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Fixed: Added Id to SELECT
                SqlCommand cmd = new SqlCommand("SELECT Id, ClientName FROM Client", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var client = new Client
                    {
                        Id = Convert.ToInt32(reader["Id"]),  // Added Id
                        ClientName = reader["ClientName"].ToString(),
                    };
                    clients.Add(client);
                }
            }
            return clients;
        }

        // Get client by Id
        public Client GetClientById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT Id, ClientName FROM Client WHERE Id = @Id", conn);
                cmd.Parameters.AddWithValue("@Id", id);
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    return new Client
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        ClientName = reader["ClientName"].ToString()
                    };
                }
            }
            return null;
        }

        // Add a new client
        public void AddClient(Client client)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO Client (ClientName) VALUES (@Name)",
                    conn
                );
                cmd.Parameters.AddWithValue("@Name", client.ClientName);
                cmd.ExecuteNonQuery();
            }
        }
    }
}