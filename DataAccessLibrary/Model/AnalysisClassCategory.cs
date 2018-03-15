using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
   public partial class AnalysisClassCategory
    {
        [Key]
        public int AnalysisClassCategoryId { get; set; }
       // public int AnalysisClassId { get; set; }
        public string Name { get; set; }
        public int Sequence { get; set; }
       // public int ProjectionMethodId { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public AnalysisClass AnalysisClass { get; set; }
        public ProjectionMethod ProjectionMethod { get; set; }
        public ICollection<RevenueExpenseData> RevenueExpenseData { get; set; }
    }
}
