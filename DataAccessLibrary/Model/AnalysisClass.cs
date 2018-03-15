using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
   public partial class AnalysisClass
    {
        [Key]
        public int AnalysisClassId { get; set; }
        public string Name { get; set; }
        //public int AnalysisId { get; set; }
        public int ClassId { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public Analysis Analysis { get; set; }
        public ICollection<AnalysisClassCategory> AnalysisClassCategory { get; set;}
    }
}
