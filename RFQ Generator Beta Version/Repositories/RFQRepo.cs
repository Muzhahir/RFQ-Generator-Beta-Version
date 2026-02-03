using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System.Repositories
{
    public class RFQRepo : BaseRepo<RFQ>
    {
        // Get all RFQs
        public List<RFQ> GetAllRFQs()
        {
            var rfqs = new List<RFQ>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Table name is already correct (RFQ)
                string query = "SELECT Id, CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode, DeliveryPoint, DeliveryTerm, Validity, Discount FROM RFQ";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            RFQ rfq = new RFQ
                            {
                                Id = reader.GetInt32(0),
                                CompanyId = reader.GetInt32(1),
                                ClientId = reader.GetInt32(2),
                                CreatedAt = reader.GetDateTime(3),
                                RFQCode = reader.GetString(4),
                                QuoteCode = reader.GetString(5),
                                DeliveryPoint = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                DeliveryTerm = reader.IsDBNull(7) ? "" : reader.GetString(7),
                                Validity = reader.IsDBNull(8) ? "" : reader.GetString(8),
                                Discount = reader.IsDBNull(9) ? 0 : reader.GetDecimal(9)
                            };
                            rfqs.Add(rfq);
                        }
                    }
                }
            }
            return rfqs;
        }

        // Get RFQ by Id
        public RFQ GetRFQById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string query = "SELECT Id, CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode, DeliveryPoint, DeliveryTerm, Validity, Discount FROM RFQ WHERE Id = @Id";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Id", id);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new RFQ
                            {
                                Id = reader.GetInt32(0),
                                CompanyId = reader.GetInt32(1),
                                ClientId = reader.GetInt32(2),
                                CreatedAt = reader.GetDateTime(3),
                                RFQCode = reader.GetString(4),
                                QuoteCode = reader.GetString(5),
                                DeliveryPoint = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                DeliveryTerm = reader.IsDBNull(7) ? "" : reader.GetString(7),
                                Validity = reader.IsDBNull(8) ? "" : reader.GetString(8),
                                Discount = reader.IsDBNull(9) ? 0 : reader.GetDecimal(9)
                            };
                        }
                    }
                }
            }
            return null;
        }

        // Add a new RFQ and return the new Id
        // Fixed: Changed return type from void to int, added SCOPE_IDENTITY()
        public int AddRFQ(RFQ rfq)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO RFQ (CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode, DeliveryPoint, DeliveryTerm, Validity, Discount) " +
                    "VALUES (@CompanyId, @ClientId, @CreatedAt, @RFQCode, @QuoteCode, @DeliveryPoint, @DeliveryTerm, @Validity, @Discount); " +
                    "SELECT SCOPE_IDENTITY();",  // Returns the new Id
                    conn
                );
                cmd.Parameters.AddWithValue("@CompanyId", rfq.CompanyId);
                cmd.Parameters.AddWithValue("@ClientId", rfq.ClientId);
                cmd.Parameters.AddWithValue("@CreatedAt", rfq.CreatedAt);
                cmd.Parameters.AddWithValue("@RFQCode", rfq.RFQCode);
                cmd.Parameters.AddWithValue("@QuoteCode", rfq.QuoteCode);
                cmd.Parameters.AddWithValue("@DeliveryPoint", string.IsNullOrEmpty(rfq.DeliveryPoint) ? (object)DBNull.Value : rfq.DeliveryPoint);
                cmd.Parameters.AddWithValue("@DeliveryTerm", string.IsNullOrEmpty(rfq.DeliveryTerm) ? (object)DBNull.Value : rfq.DeliveryTerm);
                cmd.Parameters.AddWithValue("@Validity", string.IsNullOrEmpty(rfq.Validity) ? (object)DBNull.Value : rfq.Validity);
                cmd.Parameters.AddWithValue("@Discount", rfq.Discount);

                // ExecuteScalar returns the new Id
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }
    }
}