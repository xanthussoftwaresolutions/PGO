using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
   public partial class Analysis
    {
        [Key]
        public int AnalysisId { get; set; }
        //public int CompanyId { get; set; }
        public string Name { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public Company Company { get; set; }
        public ICollection<AnalysisClass> AnalysisClass { get; set; }
    }
}
