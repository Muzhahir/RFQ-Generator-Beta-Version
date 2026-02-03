using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RFQ_Generator_System.Repositories
{
    public class TemplateRepo : BaseRepo<Template>
    {
        // Get all templates
        public List<Template> GetAllTemplates()
        {
            var templates = new List<Template>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Fixed: Changed 'Templates' to 'Template'
                SqlCommand cmd = new SqlCommand("SELECT Id, CompanyId, TemplatePath FROM Template", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var template = new Template
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        TemplatePath = reader.GetString(2)
                    };
                    templates.Add(template);
                }
            }

            // Fixed: Was returning 'new List<Template>()', now returns populated list
            return templates;
        }

        // Get template by CompanyId (1:1 relationship)
        public Template GetTemplateByCompanyId(int companyId)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT Id, CompanyId, TemplatePath FROM Template WHERE CompanyId = @CompanyId",
                    conn);
                cmd.Parameters.AddWithValue("@CompanyId", companyId);

                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return new Template
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        TemplatePath = reader.GetString(2)
                    };
                }
            }
            return null; // No template found for this company
        }

        // Get template by Id
        public Template GetTemplateById(int id)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT Id, CompanyId, TemplatePath FROM Template WHERE Id = @Id",
                    conn);
                cmd.Parameters.AddWithValue("@Id", id);

                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    return new Template
                    {
                        Id = reader.GetInt32(0),
                        CompanyId = reader.GetInt32(1),
                        TemplatePath = reader.GetString(2)
                    };
                }
            }
            return null;
        }

        // Add a template
        public void AddTemplate(Template template)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO Template (CompanyId, TemplatePath) VALUES (@CompanyId, @TemplatePath)",
                    conn
                );

                cmd.Parameters.AddWithValue("@CompanyId", template.CompanyId);
                cmd.Parameters.AddWithValue("@TemplatePath", template.TemplatePath);
                cmd.ExecuteNonQuery();
            }
        }
    }
}