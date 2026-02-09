using RFQ_Generator_System;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RFQ_Generator_System.Repositories
{
    public class CompanyRepo : BaseRepo<Company>
    {
        // Get all companies including CompanyCode
        public List<Company> GetAllCompanies()
        {
            var companies = new List<Company>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Added CompanyCode to SELECT
                SqlCommand cmd = new SqlCommand("SELECT Id, CompanyName, CompanyCode FROM Company", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var company = new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        CompanyName = reader["CompanyName"].ToString(),
                        CompanyCode = reader["CompanyCode"].ToString() // Added Mapping
                    };
                    companies.Add(company);
                }
            }

            return companies;
        }

        // Get company by Id including CompanyCode
        public Company GetCompanyById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Added CompanyCode to SELECT
                SqlCommand cmd = new SqlCommand("SELECT Id, CompanyName, CompanyCode FROM Company WHERE Id = @Id", conn);
                cmd.Parameters.AddWithValue("@Id", id);
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    return new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        CompanyName = reader["CompanyName"].ToString(),
                        CompanyCode = reader["CompanyCode"].ToString() // Added Mapping
                    };
                }
            }
            return null;
        }

        // Add a new company including CompanyCode
        public void AddCompany(Company company)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Added CompanyCode to INSERT
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO Company (CompanyName, CompanyCode) VALUES (@Name, @Code)",
                    conn
                );
                cmd.Parameters.AddWithValue("@Name", company.CompanyName);
                cmd.Parameters.AddWithValue("@Code", company.CompanyCode); // Added Parameter
                cmd.ExecuteNonQuery();
            }
        }
    }
}