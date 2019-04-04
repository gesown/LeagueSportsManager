namespace LeagueSportsManager.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class addevents : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Competition",
                c => new
                    {
                        CompetitionId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.CompetitionId);
            
            CreateTable(
                "dbo.Event",
                c => new
                    {
                        EventId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        SportId = c.Int(nullable: false),
                        ScheduleId = c.Int(nullable: false),
                        FormatId = c.Int(nullable: false),
                        CompetitionId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.EventId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Event");
            DropTable("dbo.Competition");
        }
    }
}
