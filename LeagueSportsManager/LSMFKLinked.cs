namespace LeagueSportsManager
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class LSMFKLinked : DbContext
    {
        public LSMFKLinked()
            : base("name=LSMFKLinked")
        {
        }

        public virtual DbSet<C__MigrationHistory> C__MigrationHistory { get; set; }
        public virtual DbSet<Admin> Admins { get; set; }
        public virtual DbSet<AdminType> AdminTypes { get; set; }
        public virtual DbSet<AspNetRole> AspNetRoles { get; set; }
        public virtual DbSet<AspNetUserClaim> AspNetUserClaims { get; set; }
        public virtual DbSet<AspNetUserLogin> AspNetUserLogins { get; set; }
        public virtual DbSet<AspNetUser> AspNetUsers { get; set; }
        public virtual DbSet<Contact> Contacts { get; set; }
        public virtual DbSet<Format> Formats { get; set; }
        public virtual DbSet<Ranking> Rankings { get; set; }
        public virtual DbSet<Register> Registers { get; set; }
        public virtual DbSet<Result> Results { get; set; }
        public virtual DbSet<Role> Roles { get; set; }
        public virtual DbSet<RoleType> RoleTypes { get; set; }
        public virtual DbSet<Schedule> Schedules { get; set; }
        public virtual DbSet<Score> Scores { get; set; }
        public virtual DbSet<Sport> Sports { get; set; }
        public virtual DbSet<Support> Supports { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AdminType>()
                .HasMany(e => e.Admins)
                .WithRequired(e => e.AdminType)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<AspNetRole>()
                .HasMany(e => e.AspNetUsers)
                .WithMany(e => e.AspNetRoles)
                .Map(m => m.ToTable("AspNetUserRoles").MapLeftKey("RoleId").MapRightKey("UserId"));

            modelBuilder.Entity<AspNetUser>()
                .HasMany(e => e.AspNetUserClaims)
                .WithRequired(e => e.AspNetUser)
                .HasForeignKey(e => e.UserId);

            modelBuilder.Entity<AspNetUser>()
                .HasMany(e => e.AspNetUserLogins)
                .WithRequired(e => e.AspNetUser)
                .HasForeignKey(e => e.UserId);

            modelBuilder.Entity<Contact>()
                .HasMany(e => e.Registers)
                .WithOptional(e => e.Contact)
                .HasForeignKey(e => e.Contact_ContactId);

            modelBuilder.Entity<Register>()
                .HasMany(e => e.Sports)
                .WithOptional(e => e.Register)
                .HasForeignKey(e => e.RegisterModel_RegisterId);

            modelBuilder.Entity<RoleType>()
                .HasMany(e => e.Roles)
                .WithRequired(e => e.RoleType)
                .WillCascadeOnDelete(false);
        }
    }
}
