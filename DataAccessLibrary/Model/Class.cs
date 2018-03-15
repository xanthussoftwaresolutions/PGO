using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace DataAccessLibrary.Model
{
   public partial class Class
    {
        [Key]
        public int ClassId { get; set; }
        public string Name { get; set; }
        public int ClassCode { get; set; }
    }
}
