using DataAccessLibrary;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using BusinessLibrary.ViewModel;


namespace BusinessLibrary.Repository
{
   public class BudgetRepository : IBudgetRepository
    {
        private readonly PgoDBContext pgoDBContext;
      public BudgetRepository(PgoDBContext pgoDBContext)
        {
            this.pgoDBContext = pgoDBContext;
        }
        public void add()
        {
            var listofBudget = (from a in pgoDBContext.RevenueExpenseData
                                select new RevenueExpenseDataViewModel {
                Q1=a.Q1,
                Q2=a.Q2,
                Q3=a.Q3,
                Q4=a.AnalysisClassCategory.ProjectionMethod.Name,
                Jan=a.AnalysisClassCategory.Name,
                Feb=a.AnalysisClassCategory.ProjectionMethod.Name,
                March=a.AnalysisClassCategory.Name
            });

            //PgoDBContext asd = new PgoDBContext(new Microsoft.EntityFrameworkCore.DbContextOptions<PgoDBContext>());
            var query = pgoDBContext.RevenueExpenseData.Select(x => x.AnalysisClassCategory.ProjectionMethod.Name);
        }
    }
}
