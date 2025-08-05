using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Me.Messages;
using Microsoft.Graph.Models;
using Moq;
using OutlookGraphApi.Controllers;
using Xunit;

namespace OutlookGraphApi.Tests.Controllers
{
    public class EmailControllerTests
    {
        private readonly Mock<GraphServiceClient> _mockGraphClient;
        private readonly EmailController _controller;

        public EmailControllerTests()
        {
            // Create a mock GraphServiceClient using a mock RequestAdapter
            var mockRequestAdapter = new Mock<IRequestAdapter>();
            _mockGraphClient = new Mock<GraphServiceClient>(mockRequestAdapter.Object);
            
            _controller = new EmailController(_mockGraphClient.Object);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenUnreadMessageExists_ReturnsOkWithMessageData()
        {
            // Arrange
            var expectedMessage = new Message
            {
                Subject = "Test Subject",
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "sender@example.com"
                    }
                },
                ReceivedDateTime = DateTimeOffset.Now,
                BodyPreview = "This is a test message preview"
            };

            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message> { expectedMessage }
            };

            // Mock the entire chain: Me.Messages.GetAsync()
            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            // Verify the returned object structure using dynamic
            dynamic returnValue = okResult.Value;
            Assert.Equal("Test Subject", returnValue.Subject);
            Assert.Equal("sender@example.com", returnValue.From);
            Assert.Equal(expectedMessage.ReceivedDateTime, returnValue.ReceivedDateTime);
            Assert.Equal("This is a test message preview", returnValue.BodyPreview);

            // Verify the method was called
            _mockGraphClient.Verify(
                x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()),
                Times.Once);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenNoUnreadMessages_ReturnsNotFound()
        {
            // Arrange
            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message>() // Empty list
            };

            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            Assert.IsType<NotFoundResult>(result);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenMessagesValueIsNull_ReturnsNotFound()
        {
            // Arrange
            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = null
            };

            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            Assert.IsType<NotFoundResult>(result);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenResponseIsNull_ReturnsNotFound()
        {
            // Arrange
            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync((MessageCollectionResponse)null);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            Assert.IsType<NotFoundResult>(result);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenMessageHasNullFrom_ReturnsOkWithNullFromAddress()
        {
            // Arrange
            var expectedMessage = new Message
            {
                Subject = "Test Subject",
                From = null, // Null From field
                ReceivedDateTime = DateTimeOffset.Now,
                BodyPreview = "This is a test message preview"
            };

            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message> { expectedMessage }
            };

            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            dynamic returnValue = okResult.Value;
            Assert.Null(returnValue.From);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenMessageHasNullEmailAddress_ReturnsOkWithNullFromAddress()
        {
            // Arrange
            var expectedMessage = new Message
            {
                Subject = "Test Subject",
                From = new Recipient
                {
                    EmailAddress = null // Null EmailAddress
                },
                ReceivedDateTime = DateTimeOffset.Now,
                BodyPreview = "This is a test message preview"
            };

            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message> { expectedMessage }
            };

            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            dynamic returnValue = okResult.Value;
            Assert.Null(returnValue.From);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenGraphClientThrowsException_ThrowsException()
        {
            // Arrange
            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ThrowsAsync(new ServiceException("Graph API error"));

            // Act & Assert
            await Assert.ThrowsAsync<ServiceException>(() => _controller.GetLatestUnreadAsync());
        }

        [Theory]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("Valid Subject")]
        public async Task GetLatestUnreadAsync_WithDifferentSubjects_ReturnsCorrectSubject(string subject)
        {
            // Arrange
            var expectedMessage = new Message
            {
                Subject = subject,
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "sender@example.com"
                    }
                },
                ReceivedDateTime = DateTimeOffset.Now,
                BodyPreview = "Test preview"
            };

            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message> { expectedMessage }
            };

            _mockGraphClient
                .Setup(x => x.Me.Messages.GetAsync(
                    It.IsAny<Action<Microsoft.Graph.Me.Messages.MessagesRequestBuilderGetRequestConfiguration>>(),
                    It.IsAny<CancellationToken>()))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            dynamic returnValue = okResult.Value;
            Assert.Equal(subject, returnValue.Subject);
        }
    }

    // Alternative approach: Create a wrapper interface for easier testing
    public interface IEmailService
    {
        Task<MessageCollectionResponse> GetLatestUnreadMessagesAsync();
    }

    public class EmailService : IEmailService
    {
        private readonly GraphServiceClient _graphClient;

        public EmailService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<MessageCollectionResponse> GetLatestUnreadMessagesAsync()
        {
            return await _graphClient.Me.Messages.GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Filter = "isRead eq false";
                requestConfig.QueryParameters.Top = 1;
                requestConfig.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
            });
        }
    }

    // Updated controller using the service (recommended approach)
    [ApiController]
    [Route("api/[controller]")]
    public class EmailControllerWithService : ControllerBase
    {
        private readonly IEmailService _emailService;

        public EmailControllerWithService(IEmailService emailService)
        {
            _emailService = emailService;
        }

        [HttpGet("latest-unread")]
        public async Task<IActionResult> GetLatestUnreadAsync()
        {
            var messages = await _emailService.GetLatestUnreadMessagesAsync();
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

    // Tests for the service-based controller (much cleaner)
    public class EmailControllerWithServiceTests
    {
        private readonly Mock<IEmailService> _mockEmailService;
        private readonly EmailControllerWithService _controller;

        public EmailControllerWithServiceTests()
        {
            _mockEmailService = new Mock<IEmailService>();
            _controller = new EmailControllerWithService(_mockEmailService.Object);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenUnreadMessageExists_ReturnsOkWithMessageData()
        {
            // Arrange
            var expectedMessage = new Message
            {
                Subject = "Test Subject",
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "sender@example.com"
                    }
                },
                ReceivedDateTime = DateTimeOffset.Now,
                BodyPreview = "This is a test message preview"
            };

            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message> { expectedMessage }
            };

            _mockEmailService
                .Setup(x => x.GetLatestUnreadMessagesAsync())
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            dynamic returnValue = okResult.Value;
            Assert.Equal("Test Subject", returnValue.Subject);
            Assert.Equal("sender@example.com", returnValue.From);
            Assert.Equal(expectedMessage.ReceivedDateTime, returnValue.ReceivedDateTime);
            Assert.Equal("This is a test message preview", returnValue.BodyPreview);

            _mockEmailService.Verify(x => x.GetLatestUnreadMessagesAsync(), Times.Once);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenNoUnreadMessages_ReturnsNotFound()
        {
            // Arrange
            var messageCollectionResponse = new MessageCollectionResponse
            {
                Value = new List<Message>()
            };

            _mockEmailService
                .Setup(x => x.GetLatestUnreadMessagesAsync())
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            Assert.IsType<NotFoundResult>(result);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenServiceThrowsException_ThrowsException()
        {
            // Arrange
            _mockEmailService
                .Setup(x => x.GetLatestUnreadMessagesAsync())
                .ThrowsAsync(new ServiceException("Service error"));

            // Act & Assert
            await Assert.ThrowsAsync<ServiceException>(() => _controller.GetLatestUnreadAsync());
        }
    }
}
