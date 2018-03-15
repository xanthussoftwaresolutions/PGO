using DataAccessLibrary;
using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLibrary.Repository
{
    public class PgoRepository : IPgoRepository
    {
        private readonly PgoDBContext pgoDBContext;
        PgoRepository(PgoDBContext pgoDBContext)
        {
            this.pgoDBContext = pgoDBContext;
        }
        public void add() {
                
        }
    }
}
