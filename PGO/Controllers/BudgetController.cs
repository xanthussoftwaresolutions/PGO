using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BusinessLibrary.Repository;
using Microsoft.AspNetCore.Mvc;


namespace PGO.Controllers
{
    public class BudgetController : Controller
    {
        private readonly IBudgetRepository budgetRepository; 
        public BudgetController(IBudgetRepository budgetRepository) {
            this.budgetRepository = budgetRepository;
        }
        // GET: /<controller>/
        public IActionResult CountYear()
        {
            budgetRepository.add();
            return View();
        }
    }
}
