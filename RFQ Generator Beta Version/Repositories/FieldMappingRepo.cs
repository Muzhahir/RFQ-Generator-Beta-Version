using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System.Repositories
{
    public class FieldMappingRepo : BaseRepo<FieldMapping>
    {
        // Get all field mappings
        public List<FieldMapping> GetAllFieldMappings()
        {
            var fieldMappings = new List<FieldMapping>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                // Fixed: Changed 'FieldMappings' to 'FieldMapping'
                SqlCommand cmd = new SqlCommand("SELECT * FROM FieldMapping", conn);
                SqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    var fieldMapping = new FieldMapping
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        TemplateId = Convert.ToInt32(reader["TemplateId"]),
                        FieldKey = reader["FieldKey"].ToString(),
                        ExcellCell = reader["ExcellCell"].ToString()
                    };
                    fieldMappings.Add(fieldMapping);
                }
            }
            return fieldMappings;
        }

        // Get field mappings by TemplateId (important for Excel cell mapping)
        public List<FieldMapping> GetFieldMappingsByTemplateId(int templateId)
        {
            var fieldMappings = new List<FieldMapping>();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "SELECT Id, TemplateId, FieldKey, ExcellCell FROM FieldMapping WHERE TemplateId = @TemplateId",
                    conn);
                cmd.Parameters.AddWithValue("@TemplateId", templateId);

                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    fieldMappings.Add(new FieldMapping
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        TemplateId = Convert.ToInt32(reader["TemplateId"]),
                        FieldKey = reader["FieldKey"].ToString(),
                        ExcellCell = reader["ExcellCell"].ToString()
                    });
                }
            }
            return fieldMappings;
        }

        // Add a new field mapping
        public void AddFieldMapping(FieldMapping fieldMapping)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(
                    "INSERT INTO FieldMapping (TemplateId, FieldKey, ExcellCell) VALUES (@TemplateId, @FieldKey, @ExcellCell)",
                    conn
                );
                cmd.Parameters.AddWithValue("@TemplateId", fieldMapping.TemplateId);
                cmd.Parameters.AddWithValue("@FieldKey", fieldMapping.FieldKey);
                cmd.Parameters.AddWithValue("@ExcellCell", fieldMapping.ExcellCell);
                cmd.ExecuteNonQuery();
            }
        }
    }
}