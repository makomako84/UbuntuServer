using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace UbuntuServer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class HomeController : Controller
    {
        [HttpGet("get")]
        public IActionResult Get()
        {
            return CreatedAtAction(nameof(Get), new {value = 4});
        }
    }
}