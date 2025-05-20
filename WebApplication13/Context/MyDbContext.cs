using Microsoft.EntityFrameworkCore;

namespace WebApplication13.Context
{
    public class MyDbContext : DbContext
    {
        public MyDbContext(DbContextOptions<MyDbContext> options)
           : base(options)
        {
        }
    }
}
