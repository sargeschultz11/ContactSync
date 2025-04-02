# Contributing to ContactSync

Thank you for your interest in contributing to ContactSync! This document provides guidelines and instructions for contributing to this project.

## Code of Conduct

By participating in this project, you agree to maintain a respectful and inclusive environment for everyone. Please be kind and courteous to others, and consider the impact of your words and actions.

## How Can I Contribute?

### Reporting Bugs

If you encounter a bug, please create an issue with the following information:

1. A clear, descriptive title
2. Steps to reproduce the issue
3. Expected behavior
4. Actual behavior
5. Screenshots (if applicable)
6. Your environment details (OS, PowerShell version, etc.)
7. Any additional context

### Suggesting Enhancements

If you have ideas for enhancements, please create an issue with:

1. A clear, descriptive title
2. A detailed description of the proposed enhancement
3. Any relevant examples or mockups
4. Why this enhancement would be useful to most users

### Pull Requests

We welcome pull requests for bug fixes, enhancements, and documentation improvements. Here's the process:

1. Fork the repository
2. Create a new branch for your feature
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes
6. Push to the branch
7. Open a Pull Request

### Pull Request Guidelines

- Ensure your code follows the existing style and conventions
- Update documentation if necessary
- Include tests for new features or bug fixes
- Keep pull requests focused on a single feature or fix

## Development Setup

1. Ensure you have PowerShell 5.1 or higher installed
2. For testing script functionality locally, you'll need:
   - Access to a Microsoft 365 tenant with Exchange Online
   - Appropriate Graph API permissions

## Script Conventions

When contributing code, please follow these conventions:

1. Use descriptive variable names
2. Include comments for complex logic
3. Follow the error handling patterns established in the existing code
4. Use the `Write-Log` function for logging
5. Implement appropriate throttling detection and backoff for Graph API calls

## Documentation

Documentation improvements are always welcome! If you're updating documentation:

1. Ensure examples are clear and accurate
2. Use proper Markdown formatting
3. Check for spelling and grammar errors

## Testing

Before submitting a pull request:

1. Test your changes in a non-production environment
2. Ensure your changes don't introduce performance regressions
3. Verify that existing functionality is not broken

## Feedback

If you have questions or feedback about the contribution process, please open an issue for discussion.

Thank you for contributing to ContactSync!
