using RFQ_Generator_System;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RFQ_Generator_System.Repositories
{
    public class CompanyRepo : BaseRepo<Company>
    {
        // Get all companies
        public List<Company> GetAllCompanies()
        {
            var companies = new List<Company>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Fixed: Changed 'Companies' to 'Company' and added Id to SELECT
                SqlCommand cmd = new SqlCommand("SELECT Id, CompanyName FROM Company", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var company = new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),  // Added Id
                        CompanyName = reader["CompanyName"].ToString(),
                    };
                    companies.Add(company);
                }
            }

            return companies;
        }

        // Get company by Id
        public Company GetCompanyById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("SELECT Id, CompanyName FROM Company WHERE Id = @Id", conn);
                cmd.Parameters.AddWithValue("@Id", id);
                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    return new Company
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        CompanyName = reader["CompanyName"].ToString()
                    };
                }
            }
            return null;
        }

        // Add a new company
        public void AddCompany(Company company)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Fixed: Changed 'Companies' to 'Company'
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO Company (CompanyName) VALUES (@Name)",
                    conn
                );
                cmd.Parameters.AddWithValue("@Name", company.CompanyName);
                cmd.ExecuteNonQuery();
            }
        }
    }
}