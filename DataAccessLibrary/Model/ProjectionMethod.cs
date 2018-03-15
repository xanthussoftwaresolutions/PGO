using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
  public partial  class ProjectionMethod
    {
        [Key]
        public int ProjectionMethodId { get; set; }
        public string Name { get; set; }
        public ICollection<AnalysisClassCategory> AnalysisClassCategory { get; set; }
    }
}
