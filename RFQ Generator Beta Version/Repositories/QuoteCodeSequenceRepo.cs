using System;
using System.Data.SqlClient;

namespace RFQ_Generator_System.Repositories
{
    public class QuoteCodeSequenceRepo : BaseRepo<QuoteCodeSequence>
    {
        /// <summary>
        /// Get next sequence number for a company's quote code
        /// Auto-resets to 1 if year has changed
        /// </summary>
        public int GetNextSequence(int companyId)
        {
            int currentYear = DateTime.Now.Year;
            int nextSequence = 1;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Check if sequence exists for this company
                SqlCommand cmdCheck = new SqlCommand(
                    "SELECT CurrentSequence, LastResetYear FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn);
                cmdCheck.Parameters.AddWithValue("@CompanyId", companyId);

                SqlDataReader reader = cmdCheck.ExecuteReader();

                if (reader.Read())
                {
                    int lastResetYear = Convert.ToInt32(reader["LastResetYear"]);
                    int currentSequence = Convert.ToInt32(reader["CurrentSequence"]);
                    reader.Close();

                    // Check if year has changed
                    if (lastResetYear != currentYear)
                    {
                        // Reset sequence to 1 for new year
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
                        // Increment sequence
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
                    // First time for this company - insert new record
                    SqlCommand cmdInsert = new SqlCommand(
                        "INSERT INTO QuoteCodeSequence (CompanyId, CurrentSequence, LastResetYear) VALUES (@CompanyId, 1, @Year)",
                        conn);
                    cmdInsert.Parameters.AddWithValue("@CompanyId", companyId);
                    cmdInsert.Parameters.AddWithValue("@Year", currentYear);
                    cmdInsert.ExecuteNonQuery();
                    nextSequence = 1;
                }
            }

            return nextSequence;
        }

        /// <summary>
        /// Reset all sequences to 0
        /// </summary>
        public void ResetAllSequences()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "UPDATE QuoteCodeSequence SET CurrentSequence = 0",
                    conn);
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Reset sequence for specific company
        /// </summary>
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