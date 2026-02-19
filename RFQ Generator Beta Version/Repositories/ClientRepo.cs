using RFQ_Generator_System;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class ClientRepo : BaseRepo<Client>
    {
        // Get all clients including ClientCode/Code
        public List<Client> GetAllClients()
        {
            var clients = new List<Client>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Select ClientCode as ClientCode
                SqlCommand cmd = new SqlCommand("SELECT Id, ClientName, ClientCode, DeliveryTerm FROM Client", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var client = new Client
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        ClientName = reader["ClientName"].ToString(),
                        ClientCode = reader["ClientCode"].ToString(),
                        DeliveryTerm = reader["DeliveryTerm"].ToString()
                    };
                    clients.Add(client);
                }
            }
            return clients;
        }

        // Get client by Id including ClientCode/Code
        public Client GetClientById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT Id, ClientName, ClientCode, DeliveryTerm FROM Client WHERE Id = @Id", conn);
                cmd.Parameters.AddWithValue("@Id", id);
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    return new Client
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        ClientName = reader["ClientName"].ToString(),
                        ClientCode = reader["ClientCode"].ToString(),
                        DeliveryTerm = reader["DeliveryTerm"].ToString()
                    };
                }
            }
            return null;
        }

        // Add a new client with ClientCode/Code
        public void AddClient(Client client)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO Client (ClientName, ClientCode) VALUES (@Name, @Code)",
                    conn
                );
                cmd.Parameters.AddWithValue("@Name", client.ClientName);
                cmd.Parameters.AddWithValue("@Code", client.ClientCode);
                cmd.ExecuteNonQuery();
            }
        }
    }
}