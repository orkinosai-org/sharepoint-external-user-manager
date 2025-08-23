# Developer Notes

## Library Management Approach

For comprehensive guidelines on implementing SharePoint library management functionality, refer to the **[Library Management Approach](./library-management-approach.md)** document. This document outlines:

- **Decision Process**: When to use PnPjs vs Graph API vs legacy approaches
- **Capability Matrix**: Feature support across different APIs
- **Implementation Guidance**: Code patterns and best practices
- **Fallback Strategies**: How to handle API limitations gracefully

### Quick Reference

1. **Always start with PnPjs** for SharePoint operations
2. **Use Graph API** for tenant-level operations and advanced scenarios
3. **Document fallbacks** when modern APIs are insufficient
4. **Follow the established error handling patterns** with comprehensive logging

## PR Standards

### Required Reading
- Review the [Library Management Approach](./library-management-approach.md) before contributing
- Understand the [Technical Documentation](../TECHNICAL_DOCUMENTATION.md) architecture

### Pull Request Requirements
- Ensure that all pull requests are linked to an issue
- Use the PR template which enforces our library management standards
- Each pull request should have a clear title and description
- Document which APIs were used and justify any fallback approaches
- Follow the established code review process before merging any pull request

### Implementation Standards
- **API Priority**: PnPjs → Graph API → Legacy approaches (with justification)
- **Error Handling**: Use established patterns with audit logging
- **Testing**: Cover both success and error scenarios
- **Documentation**: Update capability matrix if new limitations are discovered

For more details on the standards, please refer to the project's documentation and the comprehensive guidance in our [Library Management Approach](./library-management-approach.md) document.