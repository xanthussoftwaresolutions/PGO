using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
   public partial class Company
    {
        [Key]
        public int CompanyId { get; set; }
        public string Name { get; set; }
        //public int CompanyTypeId { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string FirstMonthOfFiscalYear { get; set; }
        public bool UseAccountNumber { get; set; }
        public string AccountNumberMask { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public CompanyType CompanyType { get; set; }
        public ICollection<Analysis> Analysis { get; set; }
    }
}
