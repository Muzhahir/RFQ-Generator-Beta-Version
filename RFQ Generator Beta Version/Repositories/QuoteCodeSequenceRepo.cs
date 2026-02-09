using System;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class QuoteCodeSequenceRepo
    {
        private readonly string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;
      Initial Catalog=RFQDB;
      Integrated Security=True";

        /// <summary>
        /// Logic for ONLY viewing what the next sequence will be.
        /// Does NOT update the database.
        /// Use this for preview when user selects company/client.
        /// </summary>
        public int PeekNextSequence(int companyId)
        {
            int currentYear = DateTime.Now.Year;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT CurrentSequence, LastResetYear FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn);
                cmd.Parameters.AddWithValue("@CompanyId", companyId);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        int lastResetYear = Convert.ToInt32(reader["LastResetYear"]);
                        int currentSequence = Convert.ToInt32(reader["CurrentSequence"]);

                        // If the year has changed, the next sequence WILL be 1
                        if (lastResetYear != currentYear)
                        {
                            return 1;
                        }

                        // Otherwise, it's just the current + 1
                        return currentSequence + 1;
                    }
                }
            }

            // If no record exists yet, the first sequence will be 1
            return 1;
        }

        /// <summary>
        /// Get next sequence number and COMMIT it to the database.
        /// Auto-resets to 1 if year has changed.
        /// Use this ONLY when actually saving the RFQ.
        /// </summary>
        public int GetNextSequence(int companyId)
        {
            int currentYear = DateTime.Now.Year;
            int nextSequence = 1;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand cmdCheck = new SqlCommand(
                    "SELECT CurrentSequence, LastResetYear FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn);
                cmdCheck.Parameters.AddWithValue("@CompanyId", companyId);

                using (SqlDataReader reader = cmdCheck.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        int lastResetYear = Convert.ToInt32(reader["LastResetYear"]);
                        int currentSequence = Convert.ToInt32(reader["CurrentSequence"]);
                        reader.Close();

                        if (lastResetYear != currentYear)
                        {
                            // COMMIT reset to 1
                            SqlCommand cmdReset = new SqlCommand(
                                "UPDATE QuoteCodeSequence SET CurrentSequence = 1, LastResetYear = @Year WHERE CompanyId = @CompanyId",
                                conn);
                            cmdReset.Parameters.AddWithValue("@Year", currentYear);
                            cmdReset.Parameters.AddWithValue("@CompanyId", companyId);
                            cmdReset.ExecuteNonQuery();
                            nextSequence = 1;
                        }
                        else
                        {
                            // COMMIT increment
                            nextSequence = currentSequence + 1;
                            SqlCommand cmdUpdate = new SqlCommand(
                                "UPDATE QuoteCodeSequence SET CurrentSequence = @Sequence WHERE CompanyId = @CompanyId",
                                conn);
                            cmdUpdate.Parameters.AddWithValue("@Sequence", nextSequence);
                            cmdUpdate.Parameters.AddWithValue("@CompanyId", companyId);
                            cmdUpdate.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        reader.Close();
                        // COMMIT initial insert
                        SqlCommand cmdInsert = new SqlCommand(
                            "INSERT INTO QuoteCodeSequence (CompanyId, CurrentSequence, LastResetYear) VALUES (@CompanyId, 1, @Year)",
                            conn);
                        cmdInsert.Parameters.AddWithValue("@CompanyId", companyId);
                        cmdInsert.Parameters.AddWithValue("@Year", currentYear);
                        cmdInsert.ExecuteNonQuery();
                        nextSequence = 1;
                    }
                }
            }

            return nextSequence;
        }

        public void ResetAllSequences()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("UPDATE QuoteCodeSequence SET CurrentSequence = 0", conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void ResetSequenceForCompany(int companyId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "UPDATE QuoteCodeSequence SET CurrentSequence = 0 WHERE CompanyId = @CompanyId",
                    conn);
                cmd.Parameters.AddWithValue("@CompanyId", companyId);
                cmd.ExecuteNonQuery();
            }
        }
    }
}