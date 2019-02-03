namespace LeagueSportsManager.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class reinit : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Admin",
                c => new
                    {
                        AdminId = c.Int(nullable: false, identity: true),
                        UserName = c.String(),
                        AdminTypeId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.AdminId);
            
            CreateTable(
                "dbo.AdminType",
                c => new
                    {
                        AdminTypeId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.AdminTypeId);
            
            CreateTable(
                "dbo.Format",
                c => new
                    {
                        FormatId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.FormatId);
            
            CreateTable(
                "dbo.Ranking",
                c => new
                    {
                        RankingId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.RankingId);
            
            CreateTable(
                "dbo.Register",
                c => new
                    {
                        RegisterId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.RegisterId);
            
            CreateTable(
                "dbo.Result",
                c => new
                    {
                        ResultId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.ResultId);
            
            CreateTable(
                "dbo.Role",
                c => new
                    {
                        RoleId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        RoleTypeId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.RoleId);
            
            CreateTable(
                "dbo.RoleType",
                c => new
                    {
                        RoleTypeId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        RoleId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.RoleTypeId);
            
            CreateTable(
                "dbo.Schedule",
                c => new
                    {
                        ScheduleId = c.Int(nullable: false, identity: true),
                        ScheduleType = c.String(),
                    })
                .PrimaryKey(t => t.ScheduleId);
            
            CreateTable(
                "dbo.Score",
                c => new
                    {
                        ScoreId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.ScoreId);
            
            CreateTable(
                "dbo.Sport",
                c => new
                    {
                        SportId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        AdminId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.SportId);
            
            CreateTable(
                "dbo.Support",
                c => new
                    {
                        SupportId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.SupportId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Support");
            DropTable("dbo.Sport");
            DropTable("dbo.Score");
            DropTable("dbo.Schedule");
            DropTable("dbo.RoleType");
            DropTable("dbo.Role");
            DropTable("dbo.Result");
            DropTable("dbo.Register");
            DropTable("dbo.Ranking");
            DropTable("dbo.Format");
            DropTable("dbo.AdminType");
            DropTable("dbo.Admin");
        }
    }
}
