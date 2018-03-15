using System;
using System.Collections.Generic;
using System.Text;

namespace DataAccessLibrary.Model
{
  public partial  class CompanyType
    {
        public int CompanyTypeId { get; set; }
        public string Description { get; set; }
        public ICollection<Company> Company { get; set; }
    }
}
