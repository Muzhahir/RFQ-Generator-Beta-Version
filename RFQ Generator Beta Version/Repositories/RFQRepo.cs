using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class RFQRepo : BaseRepo<RFQ>
    {

        public int SaveRFQ(RFQ rfq, List<RFQItem> items)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // Insert RFQ header
                        SqlCommand cmdRFQ = new SqlCommand(
                            @"INSERT INTO RFQ
                              (CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode,
                               DeliveryPoint, DeliveryTerm, Validity, Discount, Currency)
                              VALUES
                              (@CompanyId, @ClientId, @CreatedAt, @RFQCode, @QuoteCode,
                               @DeliveryPoint, @DeliveryTerm, @Validity, @Discount, @Currency);
                              SELECT CAST(SCOPE_IDENTITY() AS INT);",
                            conn,
                            transaction
                        );

                        cmdRFQ.Parameters.AddWithValue("@CompanyId", rfq.CompanyId);
                        cmdRFQ.Parameters.AddWithValue("@ClientId", rfq.ClientId);
                        cmdRFQ.Parameters.AddWithValue("@CreatedAt", rfq.CreatedAt);
                        cmdRFQ.Parameters.AddWithValue("@RFQCode", rfq.RFQCode);
                        cmdRFQ.Parameters.AddWithValue("@QuoteCode", rfq.QuoteCode);
                        cmdRFQ.Parameters.AddWithValue("@DeliveryPoint",
                            string.IsNullOrWhiteSpace(rfq.DeliveryPoint)
                                ? (object)DBNull.Value
                                : rfq.DeliveryPoint);
                        cmdRFQ.Parameters.AddWithValue("@DeliveryTerm",
                            string.IsNullOrWhiteSpace(rfq.DeliveryTerm)
                                ? (object)DBNull.Value
                                : rfq.DeliveryTerm);
                        cmdRFQ.Parameters.AddWithValue("@Validity",
                            string.IsNullOrWhiteSpace(rfq.Validity)
                                ? (object)DBNull.Value
                                : rfq.Validity);
                        cmdRFQ.Parameters.AddWithValue("@Discount", rfq.Discount);
                        cmdRFQ.Parameters.AddWithValue("@Currency",
                            string.IsNullOrWhiteSpace(rfq.Currency) ? "RM" : rfq.Currency);

                        int newRFQId = (int)cmdRFQ.ExecuteScalar();

                        // Insert RFQ items
                        foreach (var item in items)
                        {
                            SqlCommand cmdItem = new SqlCommand(
                                @"INSERT INTO RFQItem
                                  (RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName)
                                  VALUES
                                  (@RFQId, @ItemNo, @ItemDesc, @Quantity, @DeliveryTime, @UnitPrice, @UnitName)",
                                conn,
                                transaction
                            );

                            cmdItem.Parameters.AddWithValue("@RFQId", newRFQId);
                            cmdItem.Parameters.AddWithValue("@ItemNo", item.ItemNo);
                            cmdItem.Parameters.AddWithValue("@ItemDesc", item.ItemDesc);
                            cmdItem.Parameters.AddWithValue("@Quantity", item.Quantity);
                            cmdItem.Parameters.AddWithValue("@DeliveryTime", item.DeliveryTime);
                            cmdItem.Parameters.AddWithValue("@UnitPrice", item.UnitPrice);
                            cmdItem.Parameters.AddWithValue("@UnitName", item.UnitName);

                            cmdItem.ExecuteNonQuery();
                        }

                        transaction.Commit();
                        return newRFQId;
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        // -----------------------------
        // READ OPERATIONS
        // -----------------------------

        public RFQ GetRFQById(int id)
        {
            RFQ rfq = null;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"SELECT 
                    Id, CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode,
                    DeliveryPoint, DeliveryTerm, Validity, Discount, Currency
                FROM RFQ
                WHERE Id = @Id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Id", id);

                conn.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
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
                }
            }

            return rfq;
        }

        public List<RFQ> GetAllRFQs()
        {
            List<RFQ> rfqs = new List<RFQ>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string query = @"SELECT 
                    Id, CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode,
                    DeliveryPoint, DeliveryTerm, Validity, Discount, Currency
                FROM RFQ
                ORDER BY CreatedAt DESC";

                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
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
                            Currency = reader.IsDBNull(10) ? "RM" : reader.GetString(10)
                        });
                    }
                }
            }

            return rfqs;
        }


    }
}
