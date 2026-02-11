using System;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace RFQ_Generator_System.Repositories
{
    public class QuoteCodeSequenceRepo
    {
        private readonly string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;
  Initial Catalog=RFQDB;
  Integrated Security=True";

        /// <summary>
        /// Get the next sequence number for a company and INCREMENT it in the database.
        /// This should ONLY be called when actually saving an RFQ.
        /// </summary>
        public int GetNextSequence(int companyId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Check if sequence exists for this company
                SqlCommand checkCmd = new SqlCommand(
                    "SELECT CurrentSequence FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn
                );
                checkCmd.Parameters.AddWithValue("@CompanyId", companyId);

                object result = checkCmd.ExecuteScalar();

                int currentSequence;

                if (result == null)
                {
                    // No sequence exists, create one starting at 0
                    SqlCommand insertCmd = new SqlCommand(
                        "INSERT INTO QuoteCodeSequence (CompanyId, CurrentSequence, LastUpdated) " +
                        "VALUES (@CompanyId, 0, @LastUpdated)",
                        conn
                    );
                    insertCmd.Parameters.AddWithValue("@CompanyId", companyId);
                    insertCmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                    insertCmd.ExecuteNonQuery();

                    currentSequence = 0;
                }
                else
                {
                    currentSequence = Convert.ToInt32(result);
                }

                // Increment sequence for next time
                int nextSequence = currentSequence + 1;

                SqlCommand updateCmd = new SqlCommand(
                    "UPDATE QuoteCodeSequence SET CurrentSequence = @CurrentSequence, LastUpdated = @LastUpdated " +
                    "WHERE CompanyId = @CompanyId",
                    conn
                );
                updateCmd.Parameters.AddWithValue("@CurrentSequence", nextSequence);
                updateCmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                updateCmd.Parameters.AddWithValue("@CompanyId", companyId);
                updateCmd.ExecuteNonQuery();

                return currentSequence;
            }
        }

        /// <summary>
        /// Peek at what the next sequence will be WITHOUT incrementing.
        /// Used for preview display.
        /// </summary>
        public int PeekNextSequence(int companyId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand(
                    "SELECT CurrentSequence FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn
                );
                cmd.Parameters.AddWithValue("@CompanyId", companyId);

                object result = cmd.ExecuteScalar();

                if (result == null)
                {
                    // No sequence exists yet, will start at 0
                    return 0;
                }

                return Convert.ToInt32(result);
            }
        }

        /// <summary>
        /// Update sequence based on manually edited quote code.
        /// Extracts the number from the quote code and updates the sequence if it's higher.
        /// </summary>
        public void UpdateSequenceFromQuoteCode(int companyId, string quoteCode)
        {
            if (string.IsNullOrEmpty(quoteCode))
                return;

            // Extract the numeric part from the quote code
            int extractedNumber = ExtractSequenceNumber(quoteCode);

            if (extractedNumber < 0)
                return; // Could not extract valid number

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Get current sequence
                SqlCommand checkCmd = new SqlCommand(
                    "SELECT CurrentSequence FROM QuoteCodeSequence WHERE CompanyId = @CompanyId",
                    conn
                );
                checkCmd.Parameters.AddWithValue("@CompanyId", companyId);

                object result = checkCmd.ExecuteScalar();

                int currentSequence;

                if (result == null)
                {
                    // No sequence exists, create one with the extracted number
                    SqlCommand insertCmd = new SqlCommand(
                        "INSERT INTO QuoteCodeSequence (CompanyId, CurrentSequence, LastUpdated) " +
                        "VALUES (@CompanyId, @CurrentSequence, @LastUpdated)",
                        conn
                    );
                    insertCmd.Parameters.AddWithValue("@CompanyId", companyId);
                    insertCmd.Parameters.AddWithValue("@CurrentSequence", extractedNumber);
                    insertCmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                    insertCmd.ExecuteNonQuery();
                }
                else
                {
                    currentSequence = Convert.ToInt32(result);

                    // Only update if the extracted number is higher than current sequence
                    // This ensures the next quote code will be higher than the manually edited one
                    if (extractedNumber > currentSequence)
                    {
                        SqlCommand updateCmd = new SqlCommand(
                            "UPDATE QuoteCodeSequence SET CurrentSequence = @CurrentSequence, LastUpdated = @LastUpdated " +
                            "WHERE CompanyId = @CompanyId",
                            conn
                        );
                        updateCmd.Parameters.AddWithValue("@CurrentSequence", extractedNumber);
                        updateCmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                        updateCmd.Parameters.AddWithValue("@CompanyId", companyId);
                        updateCmd.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// Extract the sequence number from various quote code formats.
        /// Returns -1 if unable to extract.
        /// 
        /// Examples:
        /// CG-1224-000005 ? 5
        /// RFP-000000000123 ? 123
        /// GASB/0042/121224 ? 42
        /// QUO-ML-0099 ? 99
        /// OGIT1224-24-007 ? 7
        /// Q-000456-ABC ? 456
        /// ABC/0001234/EPOMS ? 1234
        /// ABC/QUO/SC/000789 ? 789
        /// </summary>
        private int ExtractSequenceNumber(string quoteCode)
        {
            if (string.IsNullOrEmpty(quoteCode))
                return -1;

            // Strategy: Find all sequences of digits, and pick the longest one
            // that's not clearly a date (MMDD, DDMMYY, MMyy, etc.)

            MatchCollection matches = Regex.Matches(quoteCode, @"\d+");

            int longestNonDateNumber = -1;
            int longestLength = 0;

            foreach (Match match in matches)
            {
                string numberStr = match.Value;
                int length = numberStr.Length;
                int number = int.Parse(numberStr);

                // Skip obvious date patterns
                if (length == 4 && numberStr[0] == '0' && numberStr[1] <= '1') // Likely MMDD
                    continue;
                if (length == 4 && numberStr[0] <= '1' && numberStr[1] <= '2') // Likely MMyy
                    continue;
                if (length == 6) // Likely DDMMYY or MMDDYY
                    continue;
                if (length == 2 && number <= 31) // Likely day or month or year
                    continue;

                // Keep track of the longest number that's not a date
                if (length > longestLength)
                {
                    longestLength = length;
                    longestNonDateNumber = number;
                }
            }

            return longestNonDateNumber;
        }

        /// <summary>
        /// Reset sequence for a specific company back to 0
        /// </summary>
        public void ResetSequence(int companyId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand(
                    "UPDATE QuoteCodeSequence SET CurrentSequence = 0, LastUpdated = @LastUpdated " +
                    "WHERE CompanyId = @CompanyId",
                    conn
                );
                cmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                cmd.Parameters.AddWithValue("@CompanyId", companyId);
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Reset all sequences back to 0
        /// </summary>
        public void ResetAllSequences()
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand(
                    "UPDATE QuoteCodeSequence SET CurrentSequence = 0, LastUpdated = @LastUpdated",
                    conn
                );
                cmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now);
                cmd.ExecuteNonQuery();
            }
        }
    }
}