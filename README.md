# Marketo Template Exporter

A command-line tool and library to export Marketo email templates with their HTML content and metadata.

## Background

This tool was born out of necessity during our organization's migration away from Marketo to a new marketing automation platform. We faced a significant challenge: how to efficiently export hundreds of email templates from Marketo while preserving their HTML content and structure.

While Marketo provides APIs, there weren't any readily available tools to handle bulk template exports with proper HTML formatting. This tool was created to fill that gap, making it easier for teams to:
- Export all email templates programmatically
- Preserve HTML content exactly as it appears in Marketo
- Maintain template metadata and folder structure
- Handle the process in bulk with proper error handling and progress tracking

## Original Implementation

This tool was originally part of a larger web application. If you're interested in the implementation details or need more context, you can look at the original files:

- [routes/marketo-export.js](sf-connect-app/routes/marketo-export.js) - The core implementation with all API interactions, HTML processing, and export functionality
- [views/marketo-export.ejs](sf-connect-app/views/marketo-export.ejs) - The web interface that was used to interact with the export functionality

These files show how we originally implemented the export functionality with a web interface, which might be helpful if you're looking to understand:
- How to interact with Marketo's API
- How to handle HTML template processing
- How to manage bulk exports with progress tracking
- How to structure the export data

The current CLI tool and library were extracted from this implementation to make it more portable and reusable.

## Features

- Export all email templates from your Marketo instance
- Save pristine HTML content and metadata for each template
- Create ZIP archives of exports
- Real-time progress indicators and detailed error reporting
- Maintains folder structure
- Handles authentication and token refresh automatically
- Can be used as both a CLI tool and a Node.js library

## Installation

Clone the repository:
```bash
git clone https://github.com/thedhanawada/marketo-template-exporter.git
cd marketo-template-exporter
npm install
```

To use the CLI tool globally on your machine:
```bash
npm link
```

## Configuration

Create a `.env` file with your Marketo credentials:

```env
MARKETO_CLIENT_ID=your_client_id
MARKETO_CLIENT_SECRET=your_client_secret
MARKETO_IDENTITY_URL=https://xxx-xxx-xxx.mktorest.com/identity
MARKETO_REST_URL=https://xxx-xxx-xxx.mktorest.com
```

To get these credentials:
1. Go to Marketo > Admin > Integration > LaunchPoint
2. Create a new service
3. Note down the Client ID and Secret
4. Your REST API URL can be found under Admin > Integration > Web Services

## Usage

### Command Line

```bash
# Export all templates to default directory
marketo-export export

# Export to specific directory
marketo-export export -o ./my-exports

# Export and create ZIP archive
marketo-export export -z

# Show help
marketo-export --help
```

### Programmatic Usage

```javascript
const MarketoClient = require('./lib/marketo-client');

// Initialize the client
const client = new MarketoClient({
    clientId: 'your_client_id',
    clientSecret: 'your_client_secret',
    identityUrl: 'https://xxx-xxx-xxx.mktorest.com/identity',
    restUrl: 'https://xxx-xxx-xxx.mktorest.com'
});

// Export all templates with progress tracking
const results = await client.exportAllTemplates('./output-dir', (progress) => {
    console.log(`Processed: ${progress.successful}/${progress.total}`);
});

// Export a single template
await client.exportTemplate(templateId, './output-dir');
```

## Output Structure

```
output-dir/
├── template_123/
│   ├── metadata.json    # Template metadata including folder structure
│   └── 123.html        # Clean HTML content
├── template_456/
│   ├── metadata.json
│   └── 456.html
└── marketo-templates-{timestamp}.zip (if --zip option used)
```

## Error Handling

The tool includes robust error handling:
- Automatically retries failed API calls
- Continues processing other templates if one fails
- Provides detailed error logs
- Creates error.txt files for failed templates

## Common Issues

1. **Authentication Errors**
   - Verify your Marketo credentials
   - Check API access permissions
   - Ensure your IP is whitelisted in Marketo

2. **Rate Limiting**
   - The tool automatically handles Marketo's rate limits
   - For large exports, consider using the progress callback

3. **Memory Usage**
   - For very large template sets, use the streaming export options
   - Monitor memory usage during bulk exports

## Archive Notice

This project is now archived as our Marketo migration is complete. The code is provided as-is and is no longer actively maintained. If you have questions about using this tool, feel free to reach out to me through GitHub issues. 