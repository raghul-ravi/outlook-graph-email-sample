using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;

namespace OutlookGraphApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class EmailController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;

    public EmailController(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    [HttpGet("latest-unread")]
    public async Task<IActionResult> GetLatestUnreadAsync()
    {
        var messages = await _graphClient.Me.Messages.GetAsync(requestConfig =>
        {
            requestConfig.QueryParameters.Filter = "isRead eq false";
            requestConfig.QueryParameters.Top = 1;
            requestConfig.QueryParameters.Orderby = new[] {"receivedDateTime desc"};
        });

        var message = messages?.Value?.FirstOrDefault();
        if (message == null)
        {
            return NotFound();
        }

        return Ok(new
        {
            message.Subject,
            From = message.From?.EmailAddress?.Address,
            message.ReceivedDateTime,
            message.BodyPreview
        });
    }
}
