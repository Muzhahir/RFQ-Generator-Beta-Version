using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class RFQItemRepo : BaseRepo<RFQItem>
    {
        // Get all RFQ items
        public List<RFQItem> GetAllRFQItems()
        {
            var rfqItems = new List<RFQItem>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT Id, RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName FROM RFQItem",
                    conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    rfqItems.Add(new RFQItem
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        RFQId = Convert.ToInt32(reader["RFQId"]),
                        ItemNo = Convert.ToInt32(reader["ItemNo"]),
                        ItemDesc = reader["ItemDesc"].ToString(),
                        Quantity = Convert.ToInt32(reader["Quantity"]),
                        DeliveryTime = Convert.ToInt32(reader["DeliveryTime"]),
                        UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                        UnitName = reader["UnitName"].ToString()
                    });
                }
            }

            return rfqItems;
        }

        // Get RFQ items by RFQId
        public List<RFQItem> GetRFQItemsByRFQId(int rfqId)
        {
            var rfqItems = new List<RFQItem>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT Id, RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName " +
                    "FROM RFQItem WHERE RFQId = @RFQId ORDER BY ItemNo",
                    conn);
                cmd.Parameters.AddWithValue("@RFQId", rfqId);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    rfqItems.Add(new RFQItem
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        RFQId = Convert.ToInt32(reader["RFQId"]),
                        ItemNo = Convert.ToInt32(reader["ItemNo"]),
                        ItemDesc = reader["ItemDesc"].ToString(),
                        Quantity = Convert.ToInt32(reader["Quantity"]),
                        DeliveryTime = Convert.ToInt32(reader["DeliveryTime"]),
                        UnitPrice = Convert.ToDecimal(reader["UnitPrice"]),
                        UnitName = reader["UnitName"].ToString()
                    });
                }
            }

            return rfqItems;
        }

        // Add a single RFQ item
        public void AddRFQItem(RFQItem rfqItem)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO RFQItem (RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName) " +
                    "VALUES (@RFQId, @ItemNo, @ItemDesc, @Quantity, @DeliveryTime, @UnitPrice, @UnitName)",
                    conn
                );
                cmd.Parameters.AddWithValue("@RFQId", rfqItem.RFQId);
                cmd.Parameters.AddWithValue("@ItemNo", rfqItem.ItemNo);
                cmd.Parameters.AddWithValue("@ItemDesc", rfqItem.ItemDesc);
                cmd.Parameters.AddWithValue("@Quantity", rfqItem.Quantity);
                cmd.Parameters.AddWithValue("@DeliveryTime", rfqItem.DeliveryTime);
                cmd.Parameters.AddWithValue("@UnitPrice", rfqItem.UnitPrice);
                cmd.Parameters.AddWithValue("@UnitName", rfqItem.UnitName);
                cmd.ExecuteNonQuery();
            }
        }

        // Add multiple RFQ items in one transaction (useful for bulk insert)
        public void AddRFQItems(List<RFQItem> rfqItems, SqlTransaction transaction = null)
        {
            bool ownTransaction = (transaction == null);
            SqlConnection conn = null;

            try
            {
                if (ownTransaction)
                {
                    conn = new SqlConnection(connectionString);
                    conn.Open();
                    transaction = conn.BeginTransaction();
                }

                foreach (var item in rfqItems)
                {
                    SqlCommand cmd = new SqlCommand(
                        "INSERT INTO RFQItem (RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName) " +
                        "VALUES (@RFQId, @ItemNo, @ItemDesc, @Quantity, @DeliveryTime, @UnitPrice, @UnitName)",
                        transaction.Connection,
                        transaction
                    );
                    cmd.Parameters.AddWithValue("@RFQId", item.RFQId);
                    cmd.Parameters.AddWithValue("@ItemNo", item.ItemNo);
                    cmd.Parameters.AddWithValue("@ItemDesc", item.ItemDesc);
                    cmd.Parameters.AddWithValue("@Quantity", item.Quantity);
                    cmd.Parameters.AddWithValue("@DeliveryTime", item.DeliveryTime);
                    cmd.Parameters.AddWithValue("@UnitPrice", item.UnitPrice);
                    cmd.Parameters.AddWithValue("@UnitName", item.UnitName);
                    cmd.ExecuteNonQuery();
                }

                if (ownTransaction)
                    transaction.Commit();
            }
            catch
            {
                if (ownTransaction && transaction != null)
                    transaction.Rollback();
                throw;
            }
            finally
            {
                if (ownTransaction && conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }

        // ✅ Delete all items for a given RFQ (used before re-inserting updated items)
        public void DeleteRFQItemsByRFQId(int rfqId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "DELETE FROM RFQItem WHERE RFQId = @RFQId",
                    conn
                );
                cmd.Parameters.AddWithValue("@RFQId", rfqId);
                cmd.ExecuteNonQuery();
            }
        }

        // ✅ Insert a list of items for a given RFQ (used after DeleteRFQItemsByRFQId on update)
        public void SaveRFQItems(int rfqId, List<RFQItem> items)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (var item in items)
                        {
                            SqlCommand cmd = new SqlCommand(
                                "INSERT INTO RFQItem (RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName) " +
                                "VALUES (@RFQId, @ItemNo, @ItemDesc, @Quantity, @DeliveryTime, @UnitPrice, @UnitName)",
                                conn,
                                transaction
                            );
                            cmd.Parameters.AddWithValue("@RFQId", rfqId);
                            cmd.Parameters.AddWithValue("@ItemNo", item.ItemNo);
                            cmd.Parameters.AddWithValue("@ItemDesc", item.ItemDesc);
                            cmd.Parameters.AddWithValue("@Quantity", item.Quantity);
                            cmd.Parameters.AddWithValue("@DeliveryTime", item.DeliveryTime);
                            cmd.Parameters.AddWithValue("@UnitPrice", item.UnitPrice);
                            cmd.Parameters.AddWithValue("@UnitName", item.UnitName);
                            cmd.ExecuteNonQuery();
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        internal List<RFQItem> GetItemsByRFQId(int rfqId)
        {
            throw new NotImplementedException();
        }
    }
}