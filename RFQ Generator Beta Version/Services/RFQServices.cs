using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RFQ_Generator_System.Repositories;

namespace RFQ_Generator_System.Services
{
    public class RFQService
    {
        private readonly RFQRepo rfqRepo;
        private readonly RFQItemRepo rfqItemRepo;
        private readonly TemplateRepo templateRepo;
        private readonly FieldMappingRepo fieldMappingRepo;
        private readonly CompanyRepo companyRepo;
        private readonly ClientRepo clientRepo;
        
        public RFQService()
        {
            rfqRepo = new RFQRepo();
            rfqItemRepo = new RFQItemRepo();
            templateRepo = new TemplateRepo();
            fieldMappingRepo = new FieldMappingRepo();
            companyRepo = new CompanyRepo();
            clientRepo = new ClientRepo();
        }

        /// <summary>
        /// Save RFQ header and items in a single transaction
        /// Returns the new RFQ Id
        /// </summary>
        public int SaveRFQ(RFQ rfq, List<RFQItem> items)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;
      Initial Catalog=RFQDB;
      Integrated Security=True";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // 1. Save RFQ header and get the new Id
                        SqlCommand cmdRFQ = new SqlCommand(
                            "INSERT INTO RFQ (CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode, DeliveryPoint, DeliveryTerm, Validity, Discount) " +
                            "VALUES (@CompanyId, @ClientId, @CreatedAt, @RFQCode, @QuoteCode, @DeliveryPoint, @DeliveryTerm, @Validity, @Discount); " +
                            "SELECT SCOPE_IDENTITY();",
                            conn,
                            transaction
                        );
                        cmdRFQ.Parameters.AddWithValue("@CompanyId", rfq.CompanyId);
                        cmdRFQ.Parameters.AddWithValue("@ClientId", rfq.ClientId);
                        cmdRFQ.Parameters.AddWithValue("@CreatedAt", rfq.CreatedAt);
                        cmdRFQ.Parameters.AddWithValue("@RFQCode", rfq.RFQCode);
                        cmdRFQ.Parameters.AddWithValue("@QuoteCode", rfq.QuoteCode);
                        cmdRFQ.Parameters.AddWithValue("@DeliveryPoint", string.IsNullOrEmpty(rfq.DeliveryPoint) ? (object)DBNull.Value : rfq.DeliveryPoint);
                        cmdRFQ.Parameters.AddWithValue("@DeliveryTerm", string.IsNullOrEmpty(rfq.DeliveryTerm) ? (object)DBNull.Value : rfq.DeliveryTerm);
                        cmdRFQ.Parameters.AddWithValue("@Validity", string.IsNullOrEmpty(rfq.Validity) ? (object)DBNull.Value : rfq.Validity);
                        cmdRFQ.Parameters.AddWithValue("@Discount", rfq.Discount);

                        int newRFQId = Convert.ToInt32(cmdRFQ.ExecuteScalar());

                        // 2. Save all RFQ items with the new RFQId (REMOVED DeliveryTerm from items)
                        foreach (var item in items)
                        {
                            item.RFQId = newRFQId; // Set the foreign key

                            SqlCommand cmdItem = new SqlCommand(
                                "INSERT INTO RFQItem (RFQId, ItemNo, ItemDesc, Quantity, DeliveryTime, UnitPrice, UnitName) " +
                                "VALUES (@RFQId, @ItemNo, @ItemDesc, @Quantity, @DeliveryTime, @UnitPrice, @UnitName)",
                                conn,
                                transaction
                            );
                            cmdItem.Parameters.AddWithValue("@RFQId", item.RFQId);
                            cmdItem.Parameters.AddWithValue("@ItemNo", item.ItemNo);
                            cmdItem.Parameters.AddWithValue("@ItemDesc", item.ItemDesc);
                            cmdItem.Parameters.AddWithValue("@Quantity", item.Quantity);
                            cmdItem.Parameters.AddWithValue("@DeliveryTime", item.DeliveryTime);
                            cmdItem.Parameters.AddWithValue("@UnitPrice", item.UnitPrice);
                            cmdItem.Parameters.AddWithValue("@UnitName", item.UnitName);
                            cmdItem.ExecuteNonQuery();
                        }

                        // 3. Commit transaction
                        transaction.Commit();
                        return newRFQId;
                    }
                    catch (Exception)
                    {
                        // Rollback if anything fails
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Get template based on selected CompanyId
        /// </summary>
        public Template GetTemplateByCompanyId(int companyId)
        {
            return templateRepo.GetTemplateByCompanyId(companyId);
        }

        /// <summary>
        /// Get field mappings for a template (for Excel cell mapping)
        /// </summary>
        public List<FieldMapping> GetFieldMappingsForTemplate(int templateId)
        {
            return fieldMappingRepo.GetFieldMappingsByTemplateId(templateId);
        }

        /// <summary>
        /// Get all companies (for dropdown)
        /// </summary>
        public List<Company> GetAllCompanies()
        {
            return companyRepo.GetAllCompanies();
        }

        /// <summary>
        /// Get all clients (for dropdown)
        /// </summary>
        public List<Client> GetAllClients()
        {
            return clientRepo.GetAllClients();
        }

        /// <summary>
        /// Get RFQ by Id with its items
        /// </summary>
        public (RFQ rfq, List<RFQItem> items) GetRFQWithItems(int rfqId)
        {
            var rfq = rfqRepo.GetRFQById(rfqId);
            var items = rfqItemRepo.GetRFQItemsByRFQId(rfqId);
            return (rfq, items);
        }

        /// <summary>
        /// Get all RFQs
        /// </summary>
        public List<RFQ> GetAllRFQs()
        {
            return rfqRepo.GetAllRFQs();
        }
    }
}