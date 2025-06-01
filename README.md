# Marketo Template Exporter

A command-line tool and library to export Marketo email templates with their HTML content and metadata.

## Background

This tool was born out of necessity during our organization's migration away from Marketo to a new marketing automation platform. We faced a significant challenge: how to efficiently export hundreds of email templates from Marketo while preserving their HTML content and structure.

While Marketo provides APIs, there weren't any readily available tools to handle bulk template exports with proper HTML formatting. This tool was created to fill that gap, making it easier for teams to:
- Export all email templates programmatically
- Preserve HTML content exactly as it appears in Marketo
- Maintain template metadata and folder structure
- Handle the process in bulk with proper error handling and progress tracking

## Features

- 🚀 Export all email templates from your Marketo instance
- 📄 Save pristine HTML content and metadata for each template
- 🗂️ Create ZIP archives of exports
- 📊 Real-time progress indicators and detailed error reporting
- 📁 Maintains folder structure
- 🔄 Handles authentication and token refresh automatically
- 🛠️ Can be used as both a CLI tool and a Node.js library

## Installation

```bash
npm install -g marketo-template-exporter
```

Or install locally in your project:

```bash
npm install marketo-template-exporter
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
const MarketoClient = require('marketo-template-exporter');

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

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. Here's how you can help:

- Report bugs
- Suggest new features
- Improve documentation
- Add test cases
- Enhance error handling

## License

MIT

## Author

Created by the MTC team during our platform migration project. We hope it helps others facing similar challenges! 