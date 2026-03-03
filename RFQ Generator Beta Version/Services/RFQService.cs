using System;
using System.Collections.Generic;
using System.Linq;
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

        // Cache the current quote code to avoid multiple increments
        private string cachedQuoteCode = null;
        private string cachedCompanyCode = null;
        private string cachedClientCode = null;

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

        // -----------------------------
        // Quote code generation
        // -----------------------------

        /// <summary>
        /// Preview the next quote code WITHOUT incrementing
        /// </summary>
        public string GenerateQuoteCodePreview(string companyCode, string clientCode)
        {
            if (string.IsNullOrWhiteSpace(companyCode))
                return string.Empty;

            var company = companyRepo
                .GetAllCompanies()
                .FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return string.Empty;

            int sequence = quoteCodeSequenceRepo.PeekNextSequence(company.Id);
            return FormatQuoteCode(companyCode, clientCode, sequence);
        }

        /// <summary>
        /// Generate quote code and INCREMENT sequence.
        /// Uses caching to prevent multiple increments for the same company/client.
        /// Call ClearQuoteCodeCache() when starting a new RFQ or changing client.
        /// </summary>
        public string GenerateQuoteCode(string companyCode, string clientCode)
        {
            if (string.IsNullOrWhiteSpace(companyCode))
                return string.Empty;

            // Check if we already generated a code for this company/client combination
            if (cachedQuoteCode != null &&
                cachedCompanyCode == companyCode &&
                cachedClientCode == clientCode)
            {
                return cachedQuoteCode; // Return cached code, don't increment again
            }

            var company = companyRepo
                .GetAllCompanies()
                .FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return string.Empty;

            // Only increment sequence if this is a new company/client combination
            int sequence = quoteCodeSequenceRepo.GetNextSequence(company.Id);
            string quoteCode = FormatQuoteCode(companyCode, clientCode, sequence);

            // Cache the generated code
            cachedQuoteCode = quoteCode;
            cachedCompanyCode = companyCode;
            cachedClientCode = clientCode;

            return quoteCode;
        }

        /// <summary>
        /// Clear the cached quote code. 
        /// Call this when:
        /// - User changes to a different client
        /// - User starts a new RFQ
        /// - User clicks "Reset" or "New"
        /// </summary>
        public void ClearQuoteCodeCache()
        {
            cachedQuoteCode = null;
            cachedCompanyCode = null;
            cachedClientCode = null;
        }

        /// <summary>
        /// Get the currently cached quote code without generating a new one
        /// </summary>
        public string GetCachedQuoteCode()
        {
            return cachedQuoteCode;
        }

        public void UpdateSequenceFromManualEdit(int companyId, string quoteCode)
        {
            quoteCodeSequenceRepo.UpdateSequenceFromQuoteCode(companyId, quoteCode);

            // Clear cache since user manually edited the code
            ClearQuoteCodeCache();
        }

        // -----------------------------
        // RFQ save and retrieval
        // -----------------------------

        /// <summary>
        /// Save or update RFQ header and items.
        /// - If existingId == 0: INSERT a new RFQ record. Returns new RFQ Id.
        /// - If existingId  > 0: UPDATE the existing RFQ record. Returns the same existingId.
        /// </summary>
        public int SaveRFQ(RFQ rfq, List<RFQItem> items, int existingId = 0)
        {
            int rfqId;

            if (existingId > 0)
            {
                // ✅ UPDATE existing RFQ — preserve the same quote code, don't increment
                rfq.Id = existingId;
                rfqRepo.UpdateRFQ(rfq);

                // ✅ Replace items: delete old ones and re-insert updated list
                rfqItemRepo.DeleteRFQItemsByRFQId(existingId);
                rfqItemRepo.SaveRFQItems(existingId, items);

                rfqId = existingId;
            }
            else
            {
                // ✅ INSERT new RFQ
                rfqId = rfqRepo.SaveRFQ(rfq, items);
            }

            // Clear cache after saving
            ClearQuoteCodeCache();

            return rfqId;
        }

        public (RFQ rfq, List<RFQItem> items) GetRFQWithItems(int rfqId)
        {
            var rfq = rfqRepo.GetRFQById(rfqId);
            var items = rfqItemRepo.GetRFQItemsByRFQId(rfqId);
            return (rfq, items);
        }

        public List<RFQ> GetAllRFQs()
        {
            return rfqRepo.GetAllRFQs();
        }

        // -----------------------------
        // Supporting data
        // -----------------------------

        public Template GetTemplateByCompanyId(int companyId)
        {
            return templateRepo.GetTemplateByCompanyId(companyId);
        }

        public List<FieldMapping> GetFieldMappingsForTemplate(int templateId)
        {
            return fieldMappingRepo.GetFieldMappingsByTemplateId(templateId);
        }

        public List<Company> GetAllCompanies()
        {
            return companyRepo.GetAllCompanies();
        }

        public List<Client> GetAllClients()
        {
            return clientRepo.GetAllClients();
        }

        // -----------------------------
        // Sequence management
        // -----------------------------

        public void ResetAllQuoteCodeSequences()
        {
            quoteCodeSequenceRepo.ResetAllSequences();
            ClearQuoteCodeCache();
        }

        public void ResetQuoteCodeSequence(int companyId)
        {
            quoteCodeSequenceRepo.ResetSequence(companyId);
            ClearQuoteCodeCache();
        }

        // -----------------------------
        // Private helpers
        // -----------------------------

        private string FormatQuoteCode(string companyCode, string clientCode, int sequence)
        {
            string yearShort = DateTime.Now.Year.ToString().Substring(2, 2);

            switch (companyCode)
            {
                case "CG":
                    return $"CG-{DateTime.Now:MMyy}-{sequence:D6}";
                case "DE":
                    return $"DEJSB-{sequence:D6}-{DateTime.Now:ddMMyy}";
                case "GA":
                    return $"GASB-{sequence:D4}-{DateTime.Now:ddMMyy}";
                case "MA":
                    return $"QUO-ML-{sequence:D4}";
                case "OGIT":
                    return $"OGIT{DateTime.Now:ddMM}-{yearShort}-{sequence:D3}";
                case "OP":
                    return $"Q-{sequence:D6}-{clientCode}";
                case "PO":
                    return $"PENAGA-{sequence:D7}-{clientCode}";
                case "SC":
                    return $"{clientCode}-QUO-SC-{sequence:D6}";
                default:
                    return $"{companyCode}-{sequence:D6}";
            }
        }
    }
}