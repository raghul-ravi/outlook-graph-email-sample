using Xunit;
using Moq;
using OutlookGraphApi.Controllers;
using Microsoft.Graph;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Graph.Models;

namespace OutlookGraphApi.Tests.Controllers
{
    public class EmailControllerTests
    {
        [Fact]
        public async Task GetLatestUnreadAsync_ReturnsOk_WhenUnreadMessageExists()
        {
            // Arrange
            var mockGraphClient = new Mock<GraphServiceClient>(MockBehavior.Strict, null as IAuthenticationProvider);
            var mockMessagesRequestBuilder = new Mock<MeRequestBuilder>(null, null);
            var messagesResponse = new MessageCollectionResponse
            {
                Value = new List<Message>
                {
                    new Message
                    {
                        Subject = "Test Subject",
                        From = new Recipient
                        {
                            EmailAddress = new EmailAddress { Address = "test@example.com" }
                        },
                        ReceivedDateTime = System.DateTimeOffset.UtcNow,
                        BodyPreview = "Test Body Preview"
                    }
                }
            };

            // Setup the messages endpoint
            var messagesRequestBuilder = new Mock<MessagesRequestBuilder>(null, null);
            messagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messagesResponse);

            mockMessagesRequestBuilder.SetupGet(x => x.Messages).Returns(messagesRequestBuilder.Object);
            mockGraphClient.SetupGet(x => x.Me).Returns(mockMessagesRequestBuilder.Object);

            var controller = new EmailController(mockGraphClient.Object);

            // Act
            var result = await controller.GetLatestUnreadAsync();

            // Assert
            var okResult = Assert.IsType<OkObjectResult>(result);
            var value = okResult.Value as dynamic;

            Assert.Equal("Test Subject", value.Subject);
            Assert.Equal("test@example.com", value.From);
            Assert.Equal(messagesResponse.Value.First().BodyPreview, value.BodyPreview);
        }

        [Fact]
        public async Task GetLatestUnreadAsync_ReturnsNotFound_WhenNoUnreadMessageExists()
        {
            // Arrange
            var mockGraphClient = new Mock<GraphServiceClient>(MockBehavior.Strict, null as IAuthenticationProvider);
            var mockMessagesRequestBuilder = new Mock<MeRequestBuilder>(null, null);
            var messagesResponse = new MessageCollectionResponse
            {
                Value = new List<Message>()
            };

            // Setup the messages endpoint
            var messagesRequestBuilder = new Mock<MessagesRequestBuilder>(null, null);
            messagesRequestBuilder
                .Setup(x => x.GetAsync(It.IsAny<Action<MessagesRequestBuilderGetRequestConfiguration>>(), default))
                .ReturnsAsync(messagesResponse);

            mockMessagesRequestBuilder.SetupGet(x => x.Messages).Returns(messagesRequestBuilder.Object);
            mockGraphClient.SetupGet(x => x.Me).Returns(mockMessagesRequestBuilder.Object);

            var controller = new EmailController(mockGraphClient.Object);

            // Act
            var result = await controller.GetLatestUnreadAsync();

            // Assert
            Assert.IsType<NotFoundResult>(result);
        }
    }
}
