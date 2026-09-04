using System.Diagnostics;
using FralseTech.Web.Models;
using FralseTech.Web.ViewModels;
using Microsoft.AspNetCore.Mvc;

namespace FralseTech.Web.Controllers;

public class HomeController(IConfiguration configuration) : Controller
{
    public IActionResult Index()
    {
        return View(FralseTechSiteContent.BuildLandingPage(configuration));
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
