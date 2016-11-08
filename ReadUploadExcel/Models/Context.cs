using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace ReadUploadExcel.Models
{
    public class Context : DbContext
    {
        public DbSet<Pessoa> Pessoas { get; set; }
    }
}