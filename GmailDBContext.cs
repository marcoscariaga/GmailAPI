using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;


namespace GmailAPI
{
    public class GmailDBContext : DbContext
    {
        public DbSet<Info> _dbContext { get; set; }

        public GmailDBContext() : base("Data Source=10.5.90.31;User ID=sa;Password=********;Initial Catalog=CONVISA_MAIL;Integrated Security=false;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;")
        {
        }  
    }
}
