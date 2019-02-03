using System.Web.Mvc;

namespace LeagueSportsManager.Areas.Ranking
{
    public class RankingAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "Ranking";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "Ranking_default",
                "Ranking/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}