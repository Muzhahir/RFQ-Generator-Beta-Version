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
        private readonly QuoteCodeSequenceRepo quoteCodeSequenceRepo;

        public RFQService()
        {
            rfqRepo = new RFQRepo();
            rfqItemRepo = new RFQItemRepo();
            templateRepo = new TemplateRepo();
            fieldMappingRepo = new FieldMappingRepo();
            companyRepo = new CompanyRepo();
            clientRepo = new ClientRepo();
            quoteCodeSequenceRepo = new QuoteCodeSequenceRepo();
        }

        /// <summary>
        /// Generate quote code PREVIEW based on company code and client code.
        /// Shows what the NEXT saved quote code will be (peeks at sequence without incrementing).
        /// </summary>
        public string GenerateQuoteCodePreview(string companyCode, string clientCode)
        {
            if (string.IsNullOrEmpty(companyCode))
                return "";

            // Get company to retrieve its ID for sequence
            var companies = companyRepo.GetAllCompanies();
            var company = companies.FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return "";

            // PEEK at next sequence WITHOUT incrementing
            int sequence = quoteCodeSequenceRepo.PeekNextSequence(company.Id);

            return FormatQuoteCode(companyCode, clientCode, sequence);
        }

        /// <summary>
        /// Generate quote code based on company code and client code.
        /// This COMMITS the sequence increment - use ONLY when saving.
        /// </summary>
        public string GenerateQuoteCode(string companyCode, string clientCode)
        {
            if (string.IsNullOrEmpty(companyCode))
                return "";

            // Get company to retrieve its ID for sequence
            var companies = companyRepo.GetAllCompanies();
            var company = companies.FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return "";

            // GET and INCREMENT sequence
            int sequence = quoteCodeSequenceRepo.GetNextSequence(company.Id);

            return FormatQuoteCode(companyCode, clientCode, sequence);
        }

        /// <summary>
        /// Update sequence when user manually edits quote code.
        /// This extracts the sequence number and sets it, so the next auto-generated code will be +1.
        /// </summary>
        public void UpdateSequenceFromManualEdit(int companyId, string quoteCode)
        {
            quoteCodeSequenceRepo.UpdateSequenceFromQuoteCode(companyId, quoteCode);
        }

        /// <summary>
        /// Format the quote code based on company-specific rules.
        /// All sequences now start from 0.
        /// </summary>
        private string FormatQuoteCode(string companyCode, string clientCode, int sequence)
        {
            int currentYear = DateTime.Now.Year;
            string yearShort = currentYear.ToString().Substring(2, 2); // Get last 2 digits of year
            string monthDay = DateTime.Now.ToString("MMdd");

            switch (companyCode)
            {
                case "CG":
                    // Format: CG-MMYY-NNNNNN (starts from 0)
                    return $"CG-{DateTime.Now:MMyy}-{sequence:D6}";

                case "DE":
                    // Format: RFP-NNNNNNNNNNNN (starts from 0)
                    return $"RFP-{sequence:D12}";

                case "GA":
                    // Format: GASB/NNNN/DDMMYY (starts from 0)
                    return $"GASB/{sequence:D4}/{DateTime.Now:ddMMyy}";

                case "MA":
                    // Format: QUO-ML-NNNN (starts from 0)
                    return $"QUO-ML-{sequence:D4}";

                case "OGIT":
                    // Format: OGITMMDD-YY-NNN (has year, starts from 0)
                    return $"OGIT{monthDay}-{yearShort}-{sequence:D3}";

                case "OP":
                    // Format: Q-NNNNNN-CLIENTCODE (starts from 0)
                    return $"Q-{sequence:D6}-{clientCode}";

                case "PO":
                    // Format: CLIENTCODE/NNNNNNN/EPOMS (starts from 0)
                    return $"{clientCode}/{sequence:D7}/EPOMS";

                case "SC":
                    // Format: CLIENTCODE/QUO/SC/NNNNNN (starts from 0)
                    return $"{clientCode}/QUO/SC/{sequence:D6}";

                default:
                    // Generic format (starts from 0)
                    return $"{companyCode}-{sequence:D6}";
            }
        }

        /// <summary>
        /// Save RFQ header and items in a single transaction.
        /// Returns the new RFQ Id.
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
                            "INSERT INTO RFQ (CompanyId, ClientId, CreatedAt, RFQCode, QuoteCode, DeliveryPoint, DeliveryTerm, Validity, Discount, Currency) " +
                            "VALUES (@CompanyId, @ClientId, @CreatedAt, @RFQCode, @QuoteCode, @DeliveryPoint, @DeliveryTerm, @Validity, @Discount, @Currency); " +
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
                        cmdRFQ.Parameters.AddWithValue("@Currency", string.IsNullOrEmpty(rfq.Currency) ? "RM" : rfq.Currency);

                        int newRFQId = Convert.ToInt32(cmdRFQ.ExecuteScalar());

                        // 2. Save all RFQ items with the new RFQId
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

        /// <summary>
        /// Reset all quote code sequences to 0
        /// </summary>
        public void ResetAllQuoteCodeSequences()
        {
            quoteCodeSequenceRepo.ResetAllSequences();
        }

        /// <summary>
        /// Reset quote code sequence for a specific company to 0
        /// </summary>
        public void ResetQuoteCodeSequence(int companyId)
        {
            quoteCodeSequenceRepo.ResetSequence(companyId);
        }
    }
}