using System.Web.Mvc;

namespace LeagueSportsManager.Areas.League
{
    public class LeagueAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "League";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "League_default",
                "League/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}