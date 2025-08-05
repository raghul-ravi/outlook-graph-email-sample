using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Moq;
using FluentAssertions;
using OutlookGraphApi.Controllers;

namespace OutlookGraphApi.Tests.Controllers;

public class EmailControllerTests
{
    private readonly Mock<GraphServiceClient> _mockGraphClient;
    private readonly EmailController _controller;

    public EmailControllerTests()
    {
        _mockGraphClient = new Mock<GraphServiceClient>();
        _controller = new EmailController(_mockGraphClient.Object);
    }

    [Fact]
    public async Task GetLatestUnreadAsync_WhenUnreadMessageExists_ReturnsOkWithMessage()
    {
        // Arrange
        var expectedMessage = new Message
        {
            Subject = "Test Email",
            From = new Recipient
            {
                EmailAddress = new EmailAddress { Address = "sender@example.com" }
            },
            ReceivedDateTime = DateTimeOffset.Now,
            BodyPreview = "This is a test email body"
        };

        var messageCollection = new MessageCollectionResponse
        {
            Value = new List<Message> { expectedMessage }
        };

        _mockGraphClient
            .Setup(x => x.Me.Messages.GetAsync(It.IsAny<Action<MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration>>(), default))
            .ReturnsAsync(messageCollection);

        // Act
        var result = await _controller.GetLatestUnreadAsync();

        // Assert
        result.Should().BeOfType<OkObjectResult>();
        var okResult = result as OkObjectResult;
        okResult!.StatusCode.Should().Be(200);
        
        dynamic response = okResult.Value!;
        ((string)response.Subject).Should().Be("Test Email");
        ((string)response.From).Should().Be("sender@example.com");
        ((string)response.BodyPreview).Should().Be("This is a test email body");
    }

    [Fact]
    public async Task GetLatestUnreadAsync_WhenNoUnreadMessages_ReturnsNotFound()
    {
        // Arrange
        var messageCollection = new MessageCollectionResponse
        {
            Value = new List<Message>()
        };

        _mockGraphClient
            .Setup(x => x.Me.Messages.GetAsync(It.IsAny<Action<MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration>>(), default))
            .ReturnsAsync(messageCollection);

        // Act
        var result = await _controller.GetLatestUnreadAsync();

        // Assert
        result.Should().BeOfType<NotFoundResult>();
        var notFoundResult = result as NotFoundResult;
        notFoundResult!.StatusCode.Should().Be(404);
    }

    [Fact]
    public async Task GetLatestUnreadAsync_WhenGraphServiceReturnsNull_ReturnsNotFound()
    {
        // Arrange
        _mockGraphClient
            .Setup(x => x.Me.Messages.GetAsync(It.IsAny<Action<MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration>>(), default))
            .ReturnsAsync((MessageCollectionResponse?)null);

        // Act
        var result = await _controller.GetLatestUnreadAsync();

        // Assert
        result.Should().BeOfType<NotFoundResult>();
    }

    [Fact]
    public async Task GetLatestUnreadAsync_CallsGraphServiceWithCorrectFilters()
    {
        // Arrange
        MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration? capturedConfig = null;
        
        _mockGraphClient
            .Setup(x => x.Me.Messages.GetAsync(It.IsAny<Action<MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration>>(), default))
            .Callback<Action<MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration>, CancellationToken>((action, _) =>
            {
                capturedConfig = new MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration();
                action(capturedConfig);
            })
            .ReturnsAsync(new MessageCollectionResponse { Value = new List<Message>() });

        // Act
        await _controller.GetLatestUnreadAsync();

        // Assert
        capturedConfig.Should().NotBeNull();
        capturedConfig!.QueryParameters.Filter.Should().Be("isRead eq false");
        capturedConfig.QueryParameters.Top.Should().Be(1);
        capturedConfig.QueryParameters.Orderby.Should().ContainSingle("receivedDateTime desc");
    }
}