using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class RFQRepo
    {
        private readonly string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;
            Initial Catalog=RFQDB;
            Integrated Security=True";

        /// <summary>
        /// Get RFQ by ID - Includes Currency field
        /// </summary>
        public RFQ GetRFQById(int id)
        {
            RFQ rfq = null;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                // CRITICAL: Currency must be in the SELECT list
                string query = @"SELECT 
                    Id, 
                    CompanyId, 
                    ClientId, 
                    CreatedAt, 
                    RFQCode, 
                    QuoteCode, 
                    DeliveryPoint, 
                    DeliveryTerm, 
                    Validity, 
                    Discount,
                    Currency
                FROM RFQ 
                WHERE Id = @Id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Id", id);

                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    rfq = new RFQ
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        ClientId = reader.GetInt32(2),
                        CreatedAt = reader.GetDateTime(3),
                        RFQCode = reader.GetString(4),
                        QuoteCode = reader.GetString(5),
                        DeliveryPoint = reader.IsDBNull(6) ? null : reader.GetString(6),
                        DeliveryTerm = reader.IsDBNull(7) ? null : reader.GetString(7),
                        Validity = reader.IsDBNull(8) ? null : reader.GetString(8),
                        Discount = reader.GetDecimal(9),
                        // CRITICAL: Map Currency from database
                        Currency = reader.IsDBNull(10) ? "RM" : reader.GetString(10)
                    };
                }

                reader.Close();
            }

            return rfq;
        }

        /// <summary>
        /// Get all RFQs - Includes Currency field
        /// </summary>
        public List<RFQ> GetAllRFQs()
        {
            List<RFQ> rfqs = new List<RFQ>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                // CRITICAL: Currency must be in the SELECT list
                string query = @"SELECT 
                    Id, 
                    CompanyId, 
                    ClientId, 
                    CreatedAt, 
                    RFQCode, 
                    QuoteCode, 
                    DeliveryPoint, 
                    DeliveryTerm, 
                    Validity, 
                    Discount,
                    Currency
                FROM RFQ 
                ORDER BY CreatedAt DESC";

                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    rfqs.Add(new RFQ
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        ClientId = reader.GetInt32(2),
                        CreatedAt = reader.GetDateTime(3),
                        RFQCode = reader.GetString(4),
                        QuoteCode = reader.GetString(5),
                        DeliveryPoint = reader.IsDBNull(6) ? null : reader.GetString(6),
                        DeliveryTerm = reader.IsDBNull(7) ? null : reader.GetString(7),
                        Validity = reader.IsDBNull(8) ? null : reader.GetString(8),
                        Discount = reader.GetDecimal(9),
                        // CRITICAL: Map Currency from database
                        Currency = reader.IsDBNull(10) ? "RM" : reader.GetString(10)
                    });
                }

                reader.Close();
            }

            return rfqs;
        }

        /// <summary>
        /// Check if RFQ with given code already exists
        /// </summary>
        public bool RFQCodeExists(string rfqCode)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "SELECT COUNT(*) FROM RFQ WHERE RFQCode = @RFQCode";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@RFQCode", rfqCode);

                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        /// <summary>
        /// Get RFQ by RFQ Code
        /// </summary>
        public RFQ GetRFQByCode(string rfqCode)
        {
            RFQ rfq = null;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"SELECT 
                    Id, 
                    CompanyId, 
                    ClientId, 
                    CreatedAt, 
                    RFQCode, 
                    QuoteCode, 
                    DeliveryPoint, 
                    DeliveryTerm, 
                    Validity, 
                    Discount,
                    Currency
                FROM RFQ 
                WHERE RFQCode = @RFQCode";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@RFQCode", rfqCode);

                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    rfq = new RFQ
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        ClientId = reader.GetInt32(2),
                        CreatedAt = reader.GetDateTime(3),
                        RFQCode = reader.GetString(4),
                        QuoteCode = reader.GetString(5),
                        DeliveryPoint = reader.IsDBNull(6) ? null : reader.GetString(6),
                        DeliveryTerm = reader.IsDBNull(7) ? null : reader.GetString(7),
                        Validity = reader.IsDBNull(8) ? null : reader.GetString(8),
                        Discount = reader.GetDecimal(9),
                        Currency = reader.IsDBNull(10) ? "RM" : reader.GetString(10)
                    };
                }

                reader.Close();
            }

            return rfq;
        }

        /// <summary>
        /// Update an existing RFQ
        /// </summary>
        public void UpdateRFQ(RFQ rfq)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"UPDATE RFQ SET 
                    CompanyId = @CompanyId,
                    ClientId = @ClientId,
                    CreatedAt = @CreatedAt,
                    RFQCode = @RFQCode,
                    QuoteCode = @QuoteCode,
                    DeliveryPoint = @DeliveryPoint,
                    DeliveryTerm = @DeliveryTerm,
                    Validity = @Validity,
                    Discount = @Discount,
                    Currency = @Currency
                WHERE Id = @Id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Id", rfq.Id);
                cmd.Parameters.AddWithValue("@CompanyId", rfq.CompanyId);
                cmd.Parameters.AddWithValue("@ClientId", rfq.ClientId);
                cmd.Parameters.AddWithValue("@CreatedAt", rfq.CreatedAt);
                cmd.Parameters.AddWithValue("@RFQCode", rfq.RFQCode);
                cmd.Parameters.AddWithValue("@QuoteCode", rfq.QuoteCode);
                cmd.Parameters.AddWithValue("@DeliveryPoint", string.IsNullOrEmpty(rfq.DeliveryPoint) ? (object)DBNull.Value : rfq.DeliveryPoint);
                cmd.Parameters.AddWithValue("@DeliveryTerm", string.IsNullOrEmpty(rfq.DeliveryTerm) ? (object)DBNull.Value : rfq.DeliveryTerm);
                cmd.Parameters.AddWithValue("@Validity", string.IsNullOrEmpty(rfq.Validity) ? (object)DBNull.Value : rfq.Validity);
                cmd.Parameters.AddWithValue("@Discount", rfq.Discount);
                cmd.Parameters.AddWithValue("@Currency", string.IsNullOrEmpty(rfq.Currency) ? "RM" : rfq.Currency);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Delete RFQ by ID
        /// </summary>
        public void DeleteRFQ(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = "DELETE FROM RFQ WHERE Id = @Id";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Id", id);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }
    }
}