namespace LeagueSportsManager
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class LSMEntities : DbContext
    {
        public LSMEntities()
            : base("name=LSMEntities")
        {
        }

        public virtual DbSet<Admin> Admins { get; set; }
        public virtual DbSet<AdminType> AdminTypes { get; set; }
        public virtual DbSet<Role> Roles { get; set; }
        public virtual DbSet<RoleType> RoleTypes { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AdminType>()
                .HasMany(e => e.Admins)
                .WithRequired(e => e.AdminType)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<RoleType>()
                .HasMany(e => e.Roles)
                .WithRequired(e => e.RoleType)
                .WillCascadeOnDelete(false);
        }
    }
}
