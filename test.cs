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
        private readonly Mock<MessagesRequestBuilder> _mockMessagesRequestBuilder;
        private readonly EmailController _controller;

        public EmailControllerTests()
        {
            _mockGraphClient = new Mock<GraphServiceClient>();
            _mockMessagesRequestBuilder = new Mock<MessagesRequestBuilder>();
            
            // Setup the Me.Messages property
            _mockGraphClient.Setup(x => x.Me.Messages)
                .Returns(_mockMessagesRequestBuilder.Object);
            
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            // Verify the returned object structure
            var returnValue = okResult.Value;
            var properties = returnValue.GetType().GetProperties();
            
            Assert.Contains(properties, p => p.Name == "Subject");
            Assert.Contains(properties, p => p.Name == "From");
            Assert.Contains(properties, p => p.Name == "ReceivedDateTime");
            Assert.Contains(properties, p => p.Name == "BodyPreview");

            // Verify the Graph API was called with correct parameters
            _mockMessagesRequestBuilder.Verify(
                x => x.GetAsync(
                    It.Is<Action<MessagesRequestBuilderGetRequestConfiguration>>(config => 
                        VerifyRequestConfiguration(config)),
                    default),
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
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
            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            // Use reflection to check the From property is null
            var fromProperty = okResult.Value.GetType().GetProperty("From");
            var fromValue = fromProperty?.GetValue(okResult.Value);
            Assert.Null(fromValue);
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            Assert.NotNull(okResult.Value);
            
            // Use reflection to check the From property is null
            var fromProperty = okResult.Value.GetType().GetProperty("From");
            var fromValue = fromProperty?.GetValue(okResult.Value);
            Assert.Null(fromValue);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_WhenGraphClientThrowsException_ThrowsException()
        {
            // Arrange
            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
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

            _mockMessagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messageCollectionResponse);

            // Act
            var result = await _controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            var subjectProperty = okResult.Value.GetType().GetProperty("Subject");
            var actualSubject = subjectProperty?.GetValue(okResult.Value) as string;
            Assert.Equal(subject, actualSubject);
        }

        private static bool VerifyRequestConfiguration(Action<MessagesRequestBuilderGetRequestConfiguration> configAction)
        {
            var config = new MessagesRequestBuilderGetRequestConfiguration();
            configAction(config);

            // Verify the filter is set correctly
            if (config.QueryParameters.Filter != "isRead eq false")
                return false;

            // Verify Top is set to 1
            if (config.QueryParameters.Top != 1)
                return false;

            // Verify OrderBy is set correctly
            if (config.QueryParameters.Orderby == null || 
                config.QueryParameters.Orderby.Length != 1 || 
                config.QueryParameters.Orderby[0] != "receivedDateTime desc")
                return false;

            return true;
        }
    }
}
