using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Security.Policy;
using LeagueSportsManager.Areas.Admin.Models;
using LeagueSportsManager.Areas.Competition.Models;
using LeagueSportsManager.Areas.Event.Models;
using LeagueSportsManager.Areas.Format.Models;
using LeagueSportsManager.Areas.Role.Models;
using LeagueSportsManager.Areas.Sport.Models;
using LeagueSportsManager.Areas.Support.Models;
using Microsoft.AspNet.Identity;

namespace LeagueSportsManager.Migrations
{
    internal sealed class Configuration : DbMigrationsConfiguration<LeagueSportsManager>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(LeagueSportsManager context)
        {
            try
            {


                //  This method will be called after migrating to the latest version.
                //    InitUsers(context);
                InitAdminTypes(context);
                InitRoleTypes(context);
                InitFormats(context);
                InitSports(context);
                InitSupport(context);
                InitCompetition(context);
                InitEvent(context);
            }

            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data. E.g.
            //
            //    context.People.AddOrUpdate(
            //      p => p.FullName,
            //      new Person { FullName = "Andrew Peters" },
            //      new Person { FullName = "Brice Lambson" },
            //      new Person { FullName = "Rowan Miller" }
            //    );
            //
        }

        private void InitEvent(LeagueSportsManager context)
        {
            var ev = new EventModel
            {
                Name = "PSTTCTueLeague", SportId=1,CompetitionId=1,EventId=1,FormatId=1,RoleUser=new Dictionary<int, int>(),ScheduleId=1,SupportId=new List<int>()
            };
            context.EventModels.AddOrUpdate(ev);
        }

        private void InitCompetition(LeagueSportsManager context)
        {
            var comps = new List<string>()
            {
                "Team",
                "Individual",
                "Combined",
                "Casual"
            };
            foreach (var comp in comps)
            {
                context.CompetitionModels.AddOrUpdate(new CompetitionModel { Name = comp });
            }
        }

        private void InitSupport(LeagueSportsManager context)
        {
            var refDict = new Dictionary<AspNetUser, object>();
            refDict.Add(new AspNetUser(), new Uri("http://www.redondowa.com"));
            var newContent = new SupportModel();
            newContent.Name = "InitContent";
            context.SupportModels.AddOrUpdate(newContent);
        }

        /*private void InitUsers(LeagueSportsManager context)
        {
            IPasswordHasher ihasher=new PasswordHasher();
            AspNetUser marty = new AspNetUser
            {
                Id = "LSMAdmin",
                Email = "marty@redondowa.com",
                PasswordHash = ihasher.HashPassword("nn69LSMAdmin!"),
                Hometown = "Federal Way, WA"
            };
            context.AspNetUsers.AddOrUpdate(marty);
            AspNetUser jake = new AspNetUser
            {
                Id = "PSTTCJake",
                Email = "jake@psttc.org",
                PasswordHash = ihasher.HashPassword("1234ABCE"),
                Hometown = "Auburn, WA"
            };
            context.AspNetUsers.AddOrUpdate(jake);
        }*/

        private void InitSports(LeagueSportsManager context)
        {
            var formats = new List<string>()
            {
                "Table Tennis",
                "Tennis",
                "Golf",
                "Softball",
                "Baseball",
                "Cycling",
                "Volleyball",
                "Paddle Ball",
                "Mountaineering",
                "Dressage",
                "Bull Riding"
            };
            foreach (var role in formats)
            {
                context.SportModels.AddOrUpdate(new SportModel { Name = role });
            }
        }

        private void InitAdminTypes(LeagueSportsManager context)
        {
            var formats = new List<string>()
            {
                "League",
                "Role",
                "Ranking",
                "Register",
                "Result",
                "Score",
                "Sport",
                "Schedule",
                "Support"
            };
            foreach (var role in formats)
            {
                context.AdminTypeModels.AddOrUpdate(new AdminTypeModel { Name = role });
            }
        }

        private void InitFormats(LeagueSportsManager context)
        {
            var formats = new List<string>()
            {
                "Ladder",
                "Round Robin",
                "Single Elimination",
                "Double Elimination",
                "Shotgun Golf",
                "Race",
                "Participation",
                "Judged",
                "Rally"
            };
            foreach (var role in formats)
            {
                context.FormatModels.AddOrUpdate(new FormatModel { Name = role });
            }
        }

        private void InitRoleTypes(LeagueSportsManager context)
        {
            var roleTypes = new List<string>()
            {
                "Social",
                "Professional",
                "Recreational",
                "Political",
                "Religious",
                "Competitive",
                "Casual"
            };
            foreach (var role in roleTypes)
            {
                context.RoleTypeModels.AddOrUpdate(new RoleTypeModel { Name = role });
            }
        }
    }
}
