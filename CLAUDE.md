# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Architecture Overview

This is a .NET 8 Web API application that integrates with Microsoft Graph API to access Outlook email functionality. The application uses Azure AD authentication with client credentials flow.

Key components:
- **Program.cs**: Entry point that configures Azure AD authentication and dependency injection for GraphServiceClient
- **EmailController.cs**: REST API controller exposing email-related endpoints
- **Authentication**: Uses Azure Identity library with ClientSecretCredential for service-to-service authentication

## Development Commands

```bash
# Build the application
dotnet build

# Run the application
dotnet run

# Run in watch mode (auto-rebuild on changes)
dotnet watch run

# Clean build artifacts
dotnet clean

# Restore NuGet packages
dotnet restore

# Publish for deployment
dotnet publish -c Release

# Run all tests
dotnet test

# Run tests with detailed output
dotnet test --logger "console;verbosity=detailed"

# Run tests with code coverage
dotnet test --collect:"XPlat Code Coverage"

# Run a specific test
dotnet test --filter "FullyQualifiedName~EmailControllerTests"
```

## Configuration

The application requires Azure AD credentials configured through environment variables or appsettings:
- `AZURE_TENANT_ID`: Azure AD tenant ID
- `AZURE_CLIENT_ID`: Application (client) ID
- `AZURE_CLIENT_SECRET`: Client secret

## API Endpoints

- `GET /api/email/latest-unread`: Retrieves the latest unread email from the authenticated user's mailbox

## Dependencies

- Microsoft.Graph (5.58.0): Official Microsoft Graph SDK
- Azure.Identity (1.11.0): Azure authentication library
- Target Framework: .NET 8.0