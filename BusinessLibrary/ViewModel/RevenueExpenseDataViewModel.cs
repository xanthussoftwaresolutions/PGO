
using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLibrary.ViewModel
{
  public  class RevenueExpenseDataViewModel
    {
        public int BudgetId { get; set; }
        public string Q1 { get; set; }
        public string Q2 { get; set; }
        public string Q3 { get; set; }
        public string Q4 { get; set; }
        public string Jan { get; set; }
        public string Feb { get; set; }
        public string March { get; set; }
        public string April { get; set; }
        public string May { get; set; }
        public string June { get; set; }
        public string July { get; set; }
        public string Aug { get; set; }
        public string Sept { get; set; }
        public string Oct { get; set; }
        public string Nov { get; set; }
        public string Dec { get; set; }
        public int year { get; set; }
        public string Total { get; set; }
        public int PeriodType { get; set; }
        public bool IsDeleted { get; set; }
    }
}
