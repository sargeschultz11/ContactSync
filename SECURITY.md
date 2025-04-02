# Security Policy

## Supported Versions

The following versions of ContactSync are currently being supported with security updates:

| Version | Supported          |
| ------- | ------------------ |
| 2.1.x   | :white_check_mark: |
| 2.0.x   | :x:                |

## Security Considerations

ContactSync interacts with Microsoft Graph API and handles user contact information. Please keep the following security considerations in mind:

1. **API Permissions**: The application requires `User.Read.All`, `Group.Read.All`, and `Contacts.ReadWrite` permissions. These are sensitive permissions that provide broad access to user data.

2. **Credential Security**: The solution uses client secrets stored in Azure Automation variables. Ensure these secrets are properly secured and rotated regularly.

3. **Data Handling**: The scripts handle user contact information. Ensure your implementation complies with your organization's data protection policies and any applicable regulations.

## Reporting a Vulnerability

If you discover a security vulnerability in ContactSync, please follow these steps:

1. **Do not disclose the vulnerability publicly** until it has been addressed.

2. **Submit a detailed report** by creating a new issue labeled "Security" in this repository. Include:
   - A clear description of the vulnerability
   - Steps to reproduce the issue
   - Potential impact of the vulnerability
   - Suggested fixes (if any)

3. **Response Time**: You can expect an initial response within 72 hours, and we'll aim to provide regular updates until the issue is resolved.

4. **Resolution Process**: Once a vulnerability is confirmed, we will:
   - Develop and test a fix
   - Release a security update
   - Credit you in the release notes (unless you prefer to remain anonymous)

## Best Practices for Implementation

To ensure secure use of these scripts in your environment:

1. **Principle of Least Privilege**: Configure the app registration with only the permissions necessary for the scripts to function.

2. **Regular Auditing**: Periodically review the Azure Automation runbook logs to detect any unusual activity.

3. **Secret Management**: Rotate the client secret at least every 90 days and immediately if there's any suspicion it may have been compromised.

4. **Updates**: Keep the scripts updated to the latest version to benefit from security improvements.

5. **Testing**: Always test updates in a non-production environment before deploying to production.

## Security Updates

Security updates will be released as new versions of the scripts. These updates will be documented in the release notes with details of the vulnerabilities addressed.

Thank you for helping keep ContactSync secure!
