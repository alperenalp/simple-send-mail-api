using Microsoft.AspNetCore.Mvc;
using SimpleSendMailAPI.Services;
using SimpleSendMailAPI.Dto;

namespace SimpleSendMailAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EmailController : ControllerBase
    {
        private readonly IEmailService _emailService;

        public EmailController(IEmailService emailService)
        {
            _emailService = emailService;
        }

        [HttpPost]
        public IActionResult SendEmail(EmailDto request)
        {
            _emailService.SendEmail(request);
            return Ok();
        }

        [HttpPost("all")]
        public ActionResult<List<string>> SendEmailToAllEmailFromExcel(string subject)
        {
            _emailService.SendEmailToAllEmailFromExcel(subject);
            return Ok();
        }
    }
}