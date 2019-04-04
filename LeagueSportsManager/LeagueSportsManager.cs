using LeagueSportsManager.Areas.Admin.Models;
using LeagueSportsManager.Areas.Competition.Models;
using LeagueSportsManager.Areas.Event.Models;
using LeagueSportsManager.Areas.Format.Models;
using LeagueSportsManager.Areas.Ranking.Models;
using LeagueSportsManager.Areas.Register.Models;
using LeagueSportsManager.Areas.Result.Models;
using LeagueSportsManager.Areas.Role.Models;
using LeagueSportsManager.Areas.Schedule.Models;
using LeagueSportsManager.Areas.Score.Models;
using LeagueSportsManager.Areas.Sport.Models;
using LeagueSportsManager.Areas.Support.Models;

namespace LeagueSportsManager
{
    using System;
    using System.Data.Entity;
    using System.Linq;

    public partial class LeagueSportsManager : DbContext
    {
        public LeagueSportsManager()
            : base("name=LeagueSportsManager")
        {
        }
        /*#region aspnet
        public virtual DbSet<AspNetRole> AspNetRoles { get; set; }
        public virtual DbSet<AspNetUserClaim> AspNetUserClaims { get; set; }
        public virtual DbSet<AspNetUserLogin> AspNetUserLogins { get; set; }
        public virtual DbSet<AspNetUser> AspNetUsers { get; set; }
        #endregion*/
        #region LSM context
        public virtual DbSet<AdminModel> AdminModels { get; set; }
        public virtual DbSet<AdminTypeModel> AdminTypeModels { get; set; } 
        public virtual DbSet<CompetitionModel> CompetitionModels { get; set; } 
        public virtual DbSet<EventModel> EventModels { get; set; } 
        public virtual DbSet<FormatModel> FormatModels { get; set; }
        public virtual DbSet<RankingModel> RankingModels { get; set; } 
        public virtual DbSet<RegisterModel> RegisterModels { get; set; } 
        public virtual DbSet<ResultModel> ResultModels { get;set; }
        public virtual DbSet<RoleModel> RoleModels { get; set; }
        public virtual DbSet<RoleTypeModel> RoleTypeModels { get; set; } 
        public virtual DbSet<ScoreModel> ScoreModels { get; set; }
        public virtual DbSet<SportModel> SportModels { get; set; }
        public virtual DbSet<ScheduleModel> ScheduleModels { get; set; }
        public virtual DbSet<SupportModel> SupportModels { get; set; } 
        #endregion
        #region creating
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            /*modelBuilder.Entity<AspNetRole>()
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

            modelBuilder.Entity<RoleModel>()
                .HasRequired(r => r.RoleType)
                .WithMany()
                .HasForeignKey(s => s.RoleTypeId);

            modelBuilder.Entity<AdminModel>()
                .HasRequired(r => r.AdminType)
                .WithMany()
                .HasForeignKey(s => s.AdminTypeId);
            modelBuilder.Entity<AdminModel>()
                .HasRequired(r => r.AspNetUser)
                .WithMany()
                .HasForeignKey(s => s.UserName);*/
        }
        #endregion
    }
}
