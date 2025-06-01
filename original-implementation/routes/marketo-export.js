const express = require('express');
const router = express.Router();
const axios = require('axios');
const puppeteer = require('puppeteer');

const TOKEN_EXPIRY_BUFFER_MS = 5 * 60 * 1000;

async function getMarketoAccessToken(req) {
    const sessionToken = req.session.marketoAccessToken;
    const sessionTokenExpiry = req.session.marketoTokenExpiry;

    if (sessionToken && sessionTokenExpiry && sessionTokenExpiry > (Date.now() + TOKEN_EXPIRY_BUFFER_MS)) {
        console.log("DEBUG (Marketo API): Using cached access token from session.");
        return sessionToken;
    }

    const clientId = process.env.MARKETO_CLIENT_ID;
    const clientSecret = process.env.MARKETO_CLIENT_SECRET;
    const identityUrl = process.env.MARKETO_IDENTITY_URL;

    if (!clientId || !clientSecret || !identityUrl) {
        throw new Error("Marketo API credentials (Client ID, Client Secret, or Identity URL) are not configured.");
    }

    const tokenRequestUrl = `${identityUrl}?grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}`;

    console.log(`DEBUG (Marketo API): Requesting NEW access token from ${identityUrl}`);

    try {
        const response = await axios.get(tokenRequestUrl);

        if (response.data && response.data.access_token && typeof response.data.expires_in === 'number') {
            const accessToken = response.data.access_token;
            const expiresInSeconds = response.data.expires_in;
            const expiresAt = Date.now() + (expiresInSeconds * 1000);

            console.log(`DEBUG (Marketo API): Successfully obtained NEW access token. Expires in ${expiresInSeconds}s.`);

            req.session.marketoAccessToken = accessToken;
            req.session.marketoTokenExpiry = expiresAt;

            return accessToken;
        } else {
            console.error("ERROR (Marketo API): Token response missing access_token or expires_in:", JSON.stringify(response.data));
            throw new Error("Marketo token response missing required fields.");
        }
    } catch (error) {
        console.error("ERROR (Marketo API): Failed to get NEW access token:", error.message, error.response?.status, error.response?.data ? JSON.stringify(error.response.data) : 'No response data');
        let errorMessage = `Failed to authenticate with Marketo API: ${error.message}`;
        if (error.response?.data?.error_description) {
            errorMessage = `Marketo Authentication Error: ${error.response.data.error_description}`;
        } else if (error.response?.status) {
            errorMessage = `Marketo Authentication HTTP Error: ${error.response.status} - ${error.response.statusText}`;
        }
        throw new Error(errorMessage);
    }
}

// Function to get a single folder's details
async function getFolderDetails(req, folderId) {
    if (!folderId) return null;

    const accessToken = await getMarketoAccessToken(req);
    const restUrl = process.env.MARKETO_REST_URL;
    const folderEndpoint = `${restUrl}/asset/v1/folder/${encodeURIComponent(folderId)}.json`;

    console.log(`DEBUG (Marketo Folder): Fetching folder details for ID ${folderId}`);

    try {
        const response = await axios.get(folderEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        if (response.data && response.data.success && response.data.result.length > 0) {
            console.log(`DEBUG (Marketo Folder): Successfully fetched details for folder ID ${folderId}`);
            return response.data.result[0];
        } else {
            console.warn(`WARN (Marketo Folder): No folder found with ID ${folderId} or unexpected response format`);
            return null;
        }
    } catch (error) {
        console.error(`ERROR (Marketo Folder): Failed to fetch folder ID ${folderId}:`, error.message);
        return null;
    }
}

// Function to build the complete folder path recursively
async function buildFolderPath(req, folderId) {
    if (!folderId) return [];

    const folderDetails = await getFolderDetails(req, folderId);
    if (!folderDetails) return [];

    // If this is a root folder (no parent), return just this folder
    if (!folderDetails.parent || !folderDetails.parent.id) {
        return [folderDetails];
    }

    // Get the parent path recursively, then add this folder
    const parentPath = await buildFolderPath(req, folderDetails.parent.id);
    return [...parentPath, folderDetails];
}

// Function to format the folder path as a string
function formatFolderPath(folderPath) {
    if (!folderPath || !Array.isArray(folderPath) || folderPath.length === 0) {
        return 'Unknown Location';
    }

    return folderPath.map(folder => folder.name).join(' > ');
}

// Enhanced template wrapper function
function enhancedMarketoEmailHtml(emailContent, emailId) {
    // Only replace Marketo variables
    let processedContent = emailContent
        .replace(/\{\{system\.viewAsWebpageLink\}\}/g, '#viewAsWebpage')
        .replace(/\{\{system\.unsubscribeLink\}\}/g, '#unsubscribe')
        .replace(/\{\{lead\.([^}]+)\}\}/g, (match, field) => `[Lead: ${field}]`)
        .replace(/\{\{my\.([^}]+)\}\}/g, (match, token) => `[Token: ${token}]`)
        .replace(/\{\{company\.([^}]+)\}\}/g, (match, field) => `[Company: ${field}]`)
        .replace(/\{\{#each ([^}]+)\}\}(.*?)\{\{\/each\}\}/gs, '[Each Loop Content]');

    // Ensure there's a proper DOCTYPE and viewport meta tag
    if (!processedContent.includes('<!DOCTYPE')) {
        processedContent = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>Email Preview</title>\n</head>\n<body>\n${processedContent}\n</body>\n</html>`;
    } else if (!processedContent.includes('viewport')) {
        // Add viewport meta tag if missing
        processedContent = processedContent.replace('<head>', '<head>\n<meta name="viewport" content="width=device-width, initial-scale=1.0">');
    }

    return processedContent;
}

// Original wrapper function kept for backward compatibility
function wrapMarketoEmailHtml(emailContent, emailId) {
    // Only replace Marketo variables, but preserve ALL other HTML/CSS
    let processedContent = emailContent
      .replace(/\{\{system\.viewAsWebpageLink\}\}/g, '#')
      .replace(/\{\{system\.unsubscribeLink\}\}/g, '#')
      .replace(/\{\{lead\.([^}]+)\}\}/g, 'Sample Value')
      .replace(/\{\{my\.([^}]+)\}\}/g, 'Sample Value');

    // DO NOT modify the HTML structure or styling in any way

    return `
      <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
      <html xmlns="http://www.w3.org/1999/xhtml">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Email Template</title>
        <style>
          /* Only add minimal reset styles that won't interfere with the original email */
          body { margin: 0; padding: 0; width: 100% !important; }
          img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }
          table { border-collapse: collapse; }
          /* No additional styling - let the original email's CSS take precedence */
        </style>
      </head>
      <body>
        ${processedContent}
      </body>
      </html>
    `;
}

// Add this new function to fetch the complete email HTML content through Marketo's API
async function getCompleteEmailContent(req, emailId) {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        throw new Error("Marketo REST API URL is not configured.");
    }
    if (!emailId || isNaN(emailId)) {
        throw new Error("Invalid Email ID provided for content fetch.");
    }

    const accessToken = await getMarketoAccessToken(req);

    // Try to get the full content first (most complete rendering)
    const fullContentEndpoint = `${restUrl}/asset/v1/email/${encodeURIComponent(emailId)}/fullContent.json`;

    console.log(`DEBUG (Marketo Email Content API): Fetching full content from ${fullContentEndpoint}`);

    try {
        const response = await axios.get(fullContentEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
        });

        if (response.data && response.data.success) {
            console.log(`DEBUG (Marketo Email Content API): Successfully fetched full content for Email ID ${emailId}.`);
            return response.data.result[0].content;
        } else {
            console.warn(`WARN (Marketo Email Content API): Full content API returned unsuccessful response for ID ${emailId}:`, JSON.stringify(response.data));
            // Fall back to content.json
            return fetchContentJsonBackup(req, emailId, accessToken, restUrl);
        }
    } catch (error) {
        console.warn(`WARN (Marketo Email Content API): Full content API failed (${error.message}), falling back to content.json`);
        return fetchContentJsonBackup(req, emailId, accessToken, restUrl);
    }
}

// Fallback function to fetch content when fullContent API fails
async function fetchContentJsonBackup(req, emailId, accessToken, restUrl) {
    // Use the content.json endpoint as backup
    const contentEndpoint = `${restUrl}/asset/v1/email/${encodeURIComponent(emailId)}/content.json`;

    console.log(`DEBUG (Marketo Email Content API): Trying backup content endpoint ${contentEndpoint}`);

    const response = await axios.get(contentEndpoint, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
    });

    if (response.data) {
        if (response.data.errors && response.data.errors.length > 0) {
            const apiError = response.data.errors[0];
            console.error(`ERROR (Marketo Email Content API): Marketo API returned error for content of ID ${emailId}:`, JSON.stringify(apiError));
            throw new Error(`Marketo API Error (${apiError.code}): ${apiError.message}`);
        }

        if (Array.isArray(response.data.result)) {
            let collectedHtml = '';
            response.data.result.forEach(section => {
                if (section.value && Array.isArray(section.value)) {
                    section.value.forEach(valItem => {
                        if (valItem.type === 'HTML' && valItem.value) {
                            collectedHtml += valItem.value + '\n\n';
                        }
                    });
                }
                if (section.contentType === 'html' && section.content) {
                    collectedHtml += section.content + '\n\n';
                }
            });

            if (collectedHtml.length > 0) {
                console.log(`DEBUG (Marketo Email Content API): Successfully extracted HTML content from sections for Email ID ${emailId}.`);
                return collectedHtml;
            } else {
                // Last resort - try the original fetchEmailHtmlContent method
                console.warn(`WARN (Marketo Email Content API): No usable HTML content found within result sections for ID ${emailId}.`);
                const originalHtml = await fetchEmailHtmlContent(req, emailId);
                return originalHtml;
            }
        }
    }

    throw new Error(`Could not retrieve HTML content for email ID ${emailId} through any available API methods.`);
}

// Keep original function for backward compatibility
async function fetchEmailHtmlContent(req, emailId) {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        throw new Error("Marketo REST API URL is not configured.");
    }
    if (!emailId || isNaN(emailId)) {
        throw new Error("Invalid Email ID provided for content fetch.");
    }

    const accessToken = await getMarketoAccessToken(req);
    const contentEndpoint = `${restUrl}/asset/v1/email/${encodeURIComponent(emailId)}/content.json`;

    console.log(`DEBUG (Marketo Email Content Helper): Fetching content from ${contentEndpoint}`);

    const response = await axios.get(contentEndpoint, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
    });

    if (response.data) {
        if (response.data.errors && response.data.errors.length > 0) {
            const apiError = response.data.errors[0];
            console.error(`ERROR (Marketo Email Content Helper): Marketo API returned error for content of ID ${emailId}:`, JSON.stringify(apiError));
            throw new Error(`Marketo API Error (${apiError.code}): ${apiError.message}`);
        }

        if (Array.isArray(response.data.result)) {
            let collectedHtml = '';
            response.data.result.forEach(section => {
                if (section.value && Array.isArray(section.value)) {
                    section.value.forEach(valItem => {
                        if (valItem.type === 'HTML' && valItem.value) {
                            collectedHtml += valItem.value + '\n\n<!-- Section Separator -->\n\n';
                        }
                    });
                }
                if (section.contentType === 'html' && section.content) {
                    collectedHtml += section.content + '\n\n<!-- Section Separator -->\n\n';
                }
            });

            if (collectedHtml.length > 0) {
                console.log(`DEBUG (Marketo Email Content Helper): Successfully extracted HTML content from sections for Email ID ${emailId}.`);
                return collectedHtml;
            } else {
                console.warn(`WARN (Marketo Email Content Helper): No usable HTML content found within result sections for ID ${emailId}.`, JSON.stringify(response.data));
                return null;
            }
        } else {
            console.warn(`WARN (Marketo Email Content Helper): Unexpected response format or missing 'result' array for content of ID ${emailId}:`, JSON.stringify(response.data));
            throw new Error(`Email with ID ${emailId} content unavailable (unexpected response format).`);
        }
    } else {
        console.warn(`WARN (Marketo Email Content Helper): Received empty response data for content of email ID ${emailId}.`);
        throw new Error(`Email with ID ${emailId} content unavailable (empty response).`);
    }
}

// Original function for getting email preview
async function getEmailFullContentPreview(req, emailId) {
    const accessToken = await getMarketoAccessToken(req);
    const restUrl = process.env.MARKETO_REST_URL;

    if (!restUrl) {
        throw new Error("Marketo REST API URL is not configured.");
    }

    const previewEndpoint = `${restUrl}/rest/asset/v1/email/${emailId}/fullContent.json`;

    console.log(`DEBUG: Fetching email preview from ${previewEndpoint}`);

    const params = {
        status: 'approved'
    };

    const response = await axios.get(previewEndpoint, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        params: params
    });

    if (!response.data || !response.data.success) {
        throw new Error("Failed to fetch email preview: " +
            (response.data?.errors?.[0]?.message || "Unknown error"));
    }

    return response.data.result[0].content;
}

router.get('/', async (req, res, next) => {
    try {
        res.render('marketo-export', { title: 'Marketo Export' });
    } catch (err) {
        console.error("Error rendering Marketo Export page:", err);
        next(err);
    }
});

router.post('/test-connection', async (req, res) => {
    try {
        await getMarketoAccessToken(req);
        res.status(200).json({ success: true, message: 'Successfully obtained Marketo API access token.' });
    } catch (error) {
        console.error("ERROR (Marketo Test Connection):", error.message);
        res.status(500).json({ success: false, message: error.message || 'An unexpected error occurred during Marketo connection test.' });
    }
});

router.get('/templates', async (req, res) => {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        return res.status(500).json({ success: false, message: "Marketo REST API URL is not configured." });
    }

    try {
        // Get pagination parameters from query string
        const page = parseInt(req.query.page) || 1;  // Default to page 1
        const pageSize = parseInt(req.query.pageSize) || 50;  // Default to 50 items per page

        // Calculate offset for Marketo API
        const offset = (page - 1) * pageSize;

        const accessToken = await getMarketoAccessToken(req);
        const templatesEndpoint = `${restUrl}/asset/v1/emails.json`;

        // Request only the current page from Marketo
        const params = {
            maxReturn: pageSize,
            offset: offset
        };

        const response = await axios.get(templatesEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            params: params
        });

        if (!response.data.success) {
            throw new Error(response.data.errors?.[0]?.message || "Failed to fetch templates");
        }

        // Get total count for pagination (we can use your count endpoint)
        const countResponse = await axios.get(`${req.protocol}://${req.get('host')}/marketo-export/template-count`, {
            headers: { 'Cookie': req.headers.cookie } // Pass session cookies to maintain auth
        });

        const totalTemplates = countResponse.data.totalTemplates || 0;
        const totalPages = Math.ceil(totalTemplates / pageSize);

        res.status(200).json({
            success: true,
            message: `Successfully fetched templates page ${page} of ${totalPages}.`,
            templates: response.data.result,
            pagination: {
                currentPage: page,
                pageSize: pageSize,
                totalItems: totalTemplates,
                totalPages: totalPages,
                hasNextPage: page < totalPages,
                hasPrevPage: page > 1
            }
        });

    } catch (error) {
        console.error("ERROR (Marketo Templates):", error.message);
        res.status(500).json({
            success: false,
            message: error.message || 'An unexpected error occurred while fetching templates.'
        });
    }
});

router.get('/email/:id', async (req, res) => {
    const emailId = req.params.id;
    const restUrl = process.env.MARKETO_REST_URL;

    if (!restUrl) {
        console.error("ERROR (Marketo Single Email): MARKETO_REST_URL is not configured.");
        return res.status(500).json({ success: false, message: "Marketo REST API URL is not configured." });
    }

    if (!emailId || isNaN(emailId) || parseInt(emailId, 10).toString() !== emailId.toString()) {
        console.warn("WARN (Marketo Single Email): Invalid or non-numeric Email ID received:", emailId);
        return res.status(400).json({ success: false, message: "Invalid or missing Email ID provided." });
    }

    console.log(`DEBUG (Marketo Single Email): Attempting to fetch email with ID: ${emailId}`);

    try {
        const accessToken = await getMarketoAccessToken(req);
        const emailEndpoint = `${restUrl}/asset/v1/email/${encodeURIComponent(emailId)}.json`;

        console.log(`DEBUG (Marketo Single Email): Fetching details from ${emailEndpoint}`);

        const response = await axios.get(emailEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
        });

        if (response.data) {
            if (response.data.errors && response.data.errors.length > 0) {
                const apiError = response.data.errors[0];
                console.error(`ERROR (Marketo Single Email): Marketo API returned error for ID ${emailId}:`, JSON.stringify(apiError));
                throw new Error(`Marketo API Error (${apiError.code}): ${apiError.message}`);
            }

            if (Array.isArray(response.data.result) && response.data.result.length > 0) {
                const emailDetails = response.data.result[0];
                console.log(`DEBUG (Marketo Single Email): Successfully fetched details for Email ID ${emailId}.`);

                // Extract folder ID if available
                const folderId = emailDetails.folder && emailDetails.folder.id ? emailDetails.folder.id : null;

                // If we have a folder ID, get the full folder path
                let folderPath = [];
                let folderPathString = '';

                if (folderId) {
                    folderPath = await buildFolderPath(req, folderId);
                    folderPathString = formatFolderPath(folderPath);
                    console.log(`DEBUG (Marketo Single Email): Built folder path for Email ID ${emailId}: ${folderPathString}`);
                }

                // Add folder path info to the response
                emailDetails.folderPath = folderPath;
                emailDetails.folderPathString = folderPathString;

                res.status(200).json({
                    success: true,
                    message: `Successfully fetched details for Email ID ${emailId}.`,
                    email: emailDetails
                });
            } else {
                console.warn(`WARN (Marketo Single Email): Email with ID ${emailId} not found or unexpected response format:`, JSON.stringify(response.data));
                res.status(404).json({
                    success: false,
                    message: `Email with ID ${emailId} not found or details unavailable.`
                });
            }
        } else {
            console.warn(`WARN (Marketo Single Email): Received empty response data for email ID ${emailId}.`);
            res.status(404).json({
                success: false,
                message: `Email with ID ${emailId} not found or details unavailable (empty response).`
            });
        }
    } catch (error) {
        console.error(`ERROR (Marketo Single Email): Failed to fetch email ID ${emailId}:`, error.message);

        let statusCode = 500;
        let displayMessage = error.message || `An unexpected error occurred while fetching email ID ${emailId}.`;

        if (error.message.includes('Marketo API Error')) {
            statusCode = 400;
        } else if (error.message.includes('authenticate')) {
            statusCode = 401;
        } else if (error.response?.status) {
            statusCode = error.response.status;
            displayMessage = error.response.data?.errors?.[0]?.message || error.response.statusText || `HTTP Error ${statusCode}`;
            if (statusCode === 404 && !error.message.includes('Marketo API Error')) {
                displayMessage = `Email with ID ${emailId} not found.`;
            }
        } else if (error.code === 'ECONNRESET' || error.code === 'ENOTFOUND') {
            statusCode = 503;
            displayMessage = `Network error connecting to Marketo: ${error.message}`;
        }

        res.status(statusCode).json({
            success: false,
            message: displayMessage
        });
    }
});
// Updated to use the new API approach
router.get('/email/:id/content', async (req, res) => {
    const emailId = req.params.id;

    if (!emailId || isNaN(emailId) || parseInt(emailId, 10).toString() !== emailId.toString()) {
        console.warn("WARN (Marketo Email Content): Invalid or non-numeric Email ID received:", emailId);
        return res.status(400).json({
            success: false,
            message: "Invalid or missing Email ID provided."
        });
    }

    try {
        const htmlContent = await getCompleteEmailContent(req, emailId);

        if (!htmlContent) {
            return res.status(404).json({
                success: true,
                message: `HTML content not found for Email ID ${emailId}.`,
                htmlContent: null
            });
        }

        res.status(200).json({
            success: true,
            message: `Successfully fetched HTML content for Email ID ${emailId}.`,
            htmlContent: htmlContent
        });

    } catch (error) {
        console.error(`ERROR (Marketo Email Content): Failed to fetch content for email ID ${emailId}:`, error.message);

        let statusCode = 500;
        let displayMessage = error.message || `An unexpected error occurred while fetching content for email ID ${emailId}.`;

        if (error.message.includes('Marketo API Error')) {
            statusCode = 400;
        } else if (error.message.includes('authenticate')) {
            statusCode = 401;
        } else if (error.response?.status) {
            statusCode = error.response.status;
            displayMessage = error.response.data?.errors?.[0]?.message || error.response.statusText || `HTTP Error ${statusCode}`;
            if (statusCode === 404 && !error.message.includes('Marketo API Error')) {
                displayMessage = `Email content for ID ${emailId} not found.`;
            }
        } else if (error.code === 'ECONNRESET' || error.code === 'ENOTFOUND') {
            statusCode = 503;
            displayMessage = `Network error connecting to Marketo: ${error.message}`;
        }

        res.status(statusCode).json({
            success: false,
            message: displayMessage
        });
    }
});

// Updated to use the new API approach
router.get('/email/:id/view', async (req, res) => {
    const emailId = req.params.id;

    if (!emailId || isNaN(emailId) || parseInt(emailId, 10).toString() !== emailId.toString()) {
        return res.status(400).send("Invalid or missing Email ID provided.");
    }

    try {
        const htmlContent = await getCompleteEmailContent(req, emailId);

        if (!htmlContent) {
            return res.status(404).send("No HTML content found for this email.");
        }

        // Enhanced processing
        const processedHtml = enhancedMarketoEmailHtml(htmlContent, emailId);

        res.setHeader('Content-Type', 'text/html');
        res.send(processedHtml);

    } catch (error) {
        console.error(`ERROR: Failed to render email: ${error.message}`);
        res.status(500).send(`Error: ${error.message}`);
    }
});

router.get('/email/:id/preview', async (req, res) => {
    const emailId = req.params.id;

    if (!emailId || isNaN(emailId)) {
        return res.status(400).send("Invalid Email ID provided.");
    }

    try {
        const htmlContent = await getCompleteEmailContent(req, emailId);

        if (!htmlContent) {
            return res.status(404).send("No HTML content found for this email.");
        }

        // Enhanced processing
        const processedHtml = enhancedMarketoEmailHtml(htmlContent, emailId);

        res.setHeader('Content-Type', 'text/html');
        res.send(processedHtml);

    } catch (error) {
        console.error(`ERROR: Failed to render email preview:`, error);

        let errorMessage = `Error: ${error.message}`;

        // Add more detailed error information
        if (error.response) {
            // The request was made and the server responded with a status code
            // that falls out of the range of 2xx
            console.error(`Status: ${error.response.status}`);
            console.error(`Headers:`, error.response.headers);
            console.error(`Data:`, error.response.data);

            errorMessage += `<br><br>Status: ${error.response.status}`;
            if (error.response.data && error.response.data.errors) {
                errorMessage += `<br>API Error: ${JSON.stringify(error.response.data.errors)}`;
            }

            if (error.response.status === 403) {
                errorMessage += "<br><br>This is a permissions issue. Please check:<br>1. Your API user has 'Read-Write Assets' permission<br>2. Your REST API URL is correct<br>3. The email ID exists and is accessible to your API user";
            }
        }

        res.status(500).send(errorMessage);
    }
});

// New route for direct download of HTML with named file
router.get('/email/:id/download', async (req, res) => {
    const emailId = req.params.id;

    if (!emailId || isNaN(emailId) || parseInt(emailId, 10).toString() !== emailId.toString()) {
        return res.status(400).json({
            success: false,
            message: "Invalid Email ID provided."
        });
    }

    try {
        // First, fetch the email details to get the name
        const accessToken = await getMarketoAccessToken(req);
        const restUrl = process.env.MARKETO_REST_URL;
        const emailEndpoint = `${restUrl}/asset/v1/email/${encodeURIComponent(emailId)}.json`;

        const emailResponse = await axios.get(emailEndpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
        });

        let emailName = `marketo_email_${emailId}`;

        // Extract the email name if available
        if (emailResponse.data &&
            emailResponse.data.success &&
            Array.isArray(emailResponse.data.result) &&
            emailResponse.data.result.length > 0 &&
            emailResponse.data.result[0].name) {

            // Get the name and sanitize it for use as a filename
            emailName = emailResponse.data.result[0].name
                .replace(/[^\w\s-]/g, '') // Remove special chars except spaces and hyphens
                .replace(/\s+/g, '_');    // Replace spaces with underscores
        }

        // Now get the HTML content
        const htmlContent = await getCompleteEmailContent(req, emailId);

        if (!htmlContent) {
            return res.status(404).json({
                success: false,
                message: `No HTML content found for Email ID ${emailId}.`
            });
        }

        // Enhanced processing
        const processedHtml = enhancedMarketoEmailHtml(htmlContent, emailId);

        // Set headers for file download with the email name
        res.setHeader('Content-Type', 'text/html');
        res.setHeader('Content-Disposition', `attachment; filename="${emailName}.html"`);
        res.send(processedHtml);

    } catch (error) {
        console.error(`ERROR (Marketo Email Download): Failed to download email ID ${emailId}:`, error.message);
        res.status(500).json({
            success: false,
            message: `Failed to download email: ${error.message}`
        });
    }
});

const fs = require('fs');
const path = require('path');
const archiver = require('archiver'); // You'll need to install this: npm install archiver



// Route for bulk download of all templates
router.get('/bulk-download', async (req, res) => {
    console.log('Starting bulk download of all templates...');

    // Create unique folder name based on timestamp
    const timestamp = new Date().toISOString().replace(/:/g, '-').replace(/\..+/, '');
    const exportDir = path.join(process.cwd(), 'exports', `export_${timestamp}`);

    // Create a response to track progress
    res.setHeader('Content-Type', 'text/html');
    res.write('<html><head><title>Bulk Download</title>');
    res.write('<meta http-equiv="refresh" content="5">'); // Auto-refresh every 5 seconds
    res.write('<style>body{font-family:sans-serif;line-height:1.5;max-width:800px;margin:0 auto;padding:20px;}' +
              '.error{color:red;} .success{color:green;} .progress-container{margin:20px 0;background:#f0f0f0;border-radius:4px;padding:2px;}' +
              '.progress-bar{background:#4CAF50;height:20px;border-radius:4px;text-align:center;color:white;line-height:20px;}</style>');
    res.write('</head><body>');
    res.write('<h1>Marketo Templates Bulk Download</h1>');
    res.write('<p>Please keep this page open. Your download will begin automatically when complete.</p>');
    res.write('<h3>Progress:</h3>');
    res.write('<div id="progress">Starting export process...</div>');

    // Ensure the directory exists
    if (!fs.existsSync(path.join(process.cwd(), 'exports'))) {
        fs.mkdirSync(path.join(process.cwd(), 'exports'), { recursive: true });
    }

    // Start the export process in the background
    bulkExportTemplates(req, exportDir, res).catch(error => {
        console.error('Error in bulk export:', error);
        try {
            res.write(`<div class="error">Error: ${error.message}</div>`);
            res.write('</body></html>');
            res.end();
        } catch (e) {
            console.error('Error sending error response:', e);
        }
    });
});

// Function to handle the bulk export process
async function bulkExportTemplates(req, exportDir, res) {
    try {
        // Create the export directory
        fs.mkdirSync(exportDir, { recursive: true });

        // Update progress
        res.write('<div>Created export directory...</div>');

        // Get all templates
        const accessToken = await getMarketoAccessToken(req);
        const restUrl = process.env.MARKETO_REST_URL;
        const templatesEndpoint = `${restUrl}/asset/v1/emails.json`;

        // Update progress
        res.write('<div>Fetching templates list...</div>');

        // Fetch all templates using offset-based pagination
        const allTemplates = [];
        const MAX_PAGES = 50; // Set higher than your actual number of pages
        let pageCount = 0;
        let offset = 0;
        const pageSize = 200; // Marketo's maximum

        do {
            pageCount++;
            if (pageCount > MAX_PAGES) {
                res.write(`<div class="error">Reached maximum page limit (${MAX_PAGES}). Some templates may be missing.</div>`);
                break;
            }

            // Update progress
            res.write(`<div>Fetching templates page ${pageCount} (offset: ${offset})...</div>`);

            const params = {
                maxReturn: pageSize,
                offset: offset
            };

            const response = await axios.get(templatesEndpoint, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                params: params
            });

            if (response.data && Array.isArray(response.data.result)) {
                const fetchedCount = response.data.result.length;
                allTemplates.push(...response.data.result);

                // Update progress
                res.write(`<div>Fetched page ${pageCount}, found ${fetchedCount} templates (total so far: ${allTemplates.length})...</div>`);

                // If we got fewer than pageSize results, we've reached the end
                if (fetchedCount < pageSize) {
                    res.write(`<div class="success">Reached end of templates list! Total templates found: ${allTemplates.length}</div>`);
                    break;
                }

                // Increase offset for next page
                offset += pageSize;

                // Add a small delay between requests to avoid rate limiting
                await new Promise(resolve => setTimeout(resolve, 500));

            } else {
                res.write(`<div class="error">Error fetching page ${pageCount}: Unexpected response format</div>`);
                console.error('Unexpected response format:', response.data);
                break;
            }
        } while (true);

        // Update progress
        res.write(`<div><strong>Completed fetching all templates. Found total of ${allTemplates.length} templates.</strong></div>`);
        res.write(`<div>Starting to process individual templates...</div>`);

        // Process each template - but do it in batches to avoid overwhelming the API
        const BATCH_SIZE = 5; // Process 5 templates at a time
        let processedCount = 0;
        let successCount = 0;
        let errorCount = 0;

        for (let i = 0; i < allTemplates.length; i += BATCH_SIZE) {
            const batch = allTemplates.slice(i, i + BATCH_SIZE);
            const batchNumber = Math.floor(i / BATCH_SIZE) + 1;
            const totalBatches = Math.ceil(allTemplates.length / BATCH_SIZE);

            res.write(`<div>Processing batch ${batchNumber}/${totalBatches} (templates ${i+1}-${Math.min(i+BATCH_SIZE, allTemplates.length)})...</div>`);

            // Process each template in the batch concurrently
            const batchResults = await Promise.allSettled(batch.map(async (template) => {
                try {
                    await processTemplate(req, template, exportDir);
                    processedCount++;
                    successCount++;

                    // Update progress
                    res.write(`<div>✅ Processed template ${processedCount}/${allTemplates.length}: ${template.name} (ID: ${template.id})</div>`);
                    return true;
                } catch (error) {
                    processedCount++;
                    errorCount++;
                    console.error(`Error processing template ${template.id} (${template.name}):`, error);
                    res.write(`<div class="error">❌ Error processing template ${template.id} (${template.name}): ${error.message}</div>`);
                    return false;
                }
            }));

            // Add a small delay between batches to avoid rate limiting
            await new Promise(resolve => setTimeout(resolve, 300));

            // Every 5 batches, add a summary
            if (batchNumber % 5 === 0 || i + BATCH_SIZE >= allTemplates.length) {
                const percentComplete = Math.round((processedCount / allTemplates.length) * 100);
                res.write(`<div class="progress-container">
                    <div class="progress-bar" style="width:${percentComplete}%">${percentComplete}%</div>
                </div>`);
                res.write(`<div><strong>Progress: ${processedCount}/${allTemplates.length} templates processed (${successCount} successful, ${errorCount} failed)</strong></div>`);
            }
        }

        // Create a zip file of the entire export
        res.write('<div>Creating ZIP archive of all templates...</div>');
        const zipPath = await createZipArchive(exportDir);

        // Provide download link
        res.write(`<div style="margin-top:20px; padding:15px; background-color:#f0f0f0; border-radius:5px;">
            <h3>Export Complete!</h3>
            <p>Processed ${processedCount} of ${allTemplates.length} templates (${successCount} successful, ${errorCount} failed).</p>
            <p><a href="/marketo-export/download-archive/${path.basename(zipPath)}" style="display:inline-block; padding:10px 15px; background-color:#4CAF50; color:white; text-decoration:none; border-radius:4px;">
                Download ZIP Archive
            </a></p>
        </div>`);

        res.write('</body></html>');
        res.end();

    } catch (error) {
        console.error('Error in bulk export process:', error);
        throw error;
    }
}

async function processTemplate(req, template, exportDir) {
    const templateId = template.id;

    // Create sanitized template name for folder
    let templateName = template.name
        .replace(/[^\w\s-]/g, '')  // Remove special chars
        .replace(/\s+/g, '_');     // Replace spaces with underscores

    // Add ID to ensure uniqueness
    const templateDir = path.join(exportDir, `${templateName}_${templateId}`);
    fs.mkdirSync(templateDir, { recursive: true });

    // Get folder path if available
    let folderPath = [];
    if (template.folder && template.folder.id) {
        folderPath = await buildFolderPath(req, template.folder.id);
    }
    const folderPathString = formatFolderPath(folderPath);

    // Create metadata file
    const metadata = {
        id: template.id,
        name: template.name,
        status: template.status,
        createdAt: template.createdAt,
        updatedAt: template.updatedAt,
        folder: template.folder,
        folderPath: folderPathString,
        subject: template.subject,
        fromName: template.fromName,
        fromEmail: template.fromEmail,
        // Add any other metadata fields you want to include
    };

    fs.writeFileSync(
        path.join(templateDir, 'metadata.json'),
        JSON.stringify(metadata, null, 2)
    );

    // Create a human-readable text version of the metadata
    let metadataText = `Template Name: ${template.name}\n`;
    metadataText += `ID: ${template.id}\n`;
    metadataText += `Status: ${template.status}\n`;
    metadataText += `Folder: ${template.folder ? template.folder.value : 'Unknown'}\n`;
    metadataText += `Folder Path: ${folderPathString}\n`;
    metadataText += `Created At: ${new Date(template.createdAt).toLocaleString()}\n`;
    metadataText += `Updated At: ${new Date(template.updatedAt).toLocaleString()}\n`;
    metadataText += `Subject: ${template.subject || 'N/A'}\n`;
    metadataText += `From Name: ${template.fromName || 'N/A'}\n`;
    metadataText += `From Email: ${template.fromEmail || 'N/A'}\n`;

    fs.writeFileSync(
        path.join(templateDir, 'metadata.txt'),
        metadataText
    );

    // Get HTML content
    console.log(`Processing HTML content for template ${templateId} (${template.name})...`);
    try {
        // Try to get HTML content
        const htmlContent = await getCompleteEmailContent(req, templateId);

        // Log the state to diagnose issues
        console.log(`Retrieved HTML content for template ${templateId}: ${htmlContent ? 'Content found' : 'No content'}`);

        if (htmlContent) {
            // Process the HTML content
            const processedHtml = enhancedMarketoEmailHtml(htmlContent, templateId);

            // Ensure the HTML file path is correct
            const htmlFilePath = path.join(templateDir, `${templateName}.html`);
            console.log(`Saving HTML to: ${htmlFilePath}`);

            // Save the HTML file
            fs.writeFileSync(htmlFilePath, processedHtml);
            console.log(`HTML file saved successfully for template ${templateId}`);
        } else {
            // If no HTML content, create a placeholder file
            fs.writeFileSync(
                path.join(templateDir, 'no_content.txt'),
                'No HTML content was found for this template.'
            );
            console.log(`No HTML content found for template ${templateId}, created placeholder file`);
        }
    } catch (error) {
        console.error(`Error processing HTML for template ${templateId}:`, error);

        // Create an error file
        fs.writeFileSync(
            path.join(templateDir, 'error.txt'),
            `Error retrieving HTML content: ${error.message}`
        );

        // Don't throw the error - just log it and continue with other templates
        console.log(`Created error file for template ${templateId}`);
        throw error; // Re-throw to properly count errors
    }

    console.log(`Completed processing template ${templateId}`);
    return true;
}



// Helper function to create a ZIP archive of the export
async function createZipArchive(exportDir) {
    return new Promise((resolve, reject) => {
        const zipFileName = `${path.basename(exportDir)}.zip`;
        const zipPath = path.join(process.cwd(), 'exports', zipFileName);

        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', {
            zlib: { level: 5 } // Compression level
        });

        output.on('close', () => {
            console.log(`Archive created: ${zipPath}, size: ${archive.pointer()} bytes`);
            resolve(zipPath);
        });

        archive.on('error', (err) => {
            reject(err);
        });

        archive.pipe(output);
        archive.directory(exportDir, false);
        archive.finalize();
    });
}

// Route to list all available exports
router.get('/exports', async (req, res) => {
    try {
        const exportsDir = path.join(process.cwd(), 'exports');
        if (!fs.existsSync(exportsDir)) {
            fs.mkdirSync(exportsDir, { recursive: true });
        }

        // Read the exports directory
        const files = fs.readdirSync(exportsDir);

        // Filter to directories and zip files
        const exports = files.filter(file => {
            const fullPath = path.join(exportsDir, file);
            return fs.statSync(fullPath).isDirectory() || file.endsWith('.zip');
        });

        res.render('exports', {
            title: 'Marketo Exports',
            exports: exports
        });
    } catch (err) {
        console.error("Error listing exports:", err);
        res.status(500).send("Error listing exports: " + err.message);
    }
});

router.get('/template-count', async (req, res) => {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        console.error("ERROR: MARKETO_REST_URL is not configured.");
        return res.status(500).json({ success: false, message: "Marketo REST API URL is not configured." });
    }

    try {
        // Check if the includeData parameter is provided
        const includeData = req.query.includeData === 'true';

        const accessToken = await getMarketoAccessToken(req);
        const templatesEndpoint = `${restUrl}/asset/v1/emails.json`;

        let totalTemplates = 0;
        const MAX_PAGES = 100; // Set a reasonable limit
        let pageCount = 0;
        let moreResults = true;
        let offset = 0;
        const maxReturn = 200; // Maximum allowed by the API

        // Only collect all templates if includeData is true
        let allTemplates = includeData ? [] : null;

        console.log("DEBUG: Starting verification of total template count...");
        console.log(`DEBUG: Include full template data: ${includeData}`);

        while (moreResults && pageCount < MAX_PAGES) {
            pageCount++;

            console.log(`DEBUG: Fetching count page ${pageCount} with offset ${offset}...`);

            const params = {
                maxReturn: maxReturn,
                offset: offset
            };

            const response = await axios.get(templatesEndpoint, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                params: params
            });

            if (response.data && Array.isArray(response.data.result)) {
                const pageTemplates = response.data.result.length;
                totalTemplates += pageTemplates;

                // Store template data if requested
                if (includeData) {
                    allTemplates.push(...response.data.result);
                }

                console.log(`DEBUG: Page ${pageCount} has ${pageTemplates} templates. Total so far: ${totalTemplates}.`);

                // If we got fewer results than maxReturn, we've reached the end
                if (pageTemplates < maxReturn) {
                    moreResults = false;
                    console.log(`DEBUG: Received fewer than ${maxReturn} results (${pageTemplates}), ending pagination.`);
                } else {
                    // Increment the offset for the next page
                    offset += maxReturn;
                }
            } else {
                console.warn("WARN: Unexpected response format:", JSON.stringify(response.data));
                moreResults = false;
            }
        }

        // Create the response object
        const responseData = {
            success: true,
            message: `Found ${totalTemplates} templates across ${pageCount} pages.`,
            totalTemplates: totalTemplates,
            pagesChecked: pageCount,
            hasMore: moreResults
        };

        // Add templates data if requested
        if (includeData) {
            responseData.templates = allTemplates;
            console.log(`DEBUG: Returning full data for ${allTemplates.length} templates.`);
        }

        return res.status(200).json(responseData);

    } catch (error) {
        console.error("ERROR: Failed to count templates:", error.message);
        return res.status(500).json({
            success: false,
            message: error.message || 'An unexpected error occurred while counting templates.'
        });
    }
});

const { Parser } = require('json2csv');

// Existing all-templates route with CSV export option
router.get('/all-templates', async (req, res) => {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        console.error("ERROR: MARKETO_REST_URL is not configured.");
        return res.status(500).json({ success: false, message: "Marketo REST API URL is not configured." });
    }

    // Check if CSV format is requested
    const format = req.query.format?.toLowerCase();

    try {
        const accessToken = await getMarketoAccessToken(req);
        const templatesEndpoint = `${restUrl}/asset/v1/emails.json`;

        let allTemplates = [];
        const MAX_PAGES = 100;
        let pageCount = 0;
        let moreResults = true;
        let offset = 0;
        const maxReturn = 200; // Maximum allowed by the API

        console.log(`DEBUG (All Templates): Starting fetch of all templates (format: ${format || 'json'})...`);

        // Show a warning about the potential size of the response
        console.log("WARNING: This endpoint returns ALL templates and may produce a very large response");

        while (moreResults && pageCount < MAX_PAGES) {
            pageCount++;

            console.log(`DEBUG (All Templates): Fetching page ${pageCount} with offset ${offset}...`);

            const params = {
                maxReturn: maxReturn,
                offset: offset
            };

            const response = await axios.get(templatesEndpoint, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                params: params
            });

            if (response.data?.success && Array.isArray(response.data.result)) {
                const pageTemplates = response.data.result.length;
                allTemplates.push(...response.data.result);

                console.log(`DEBUG (All Templates): Page ${pageCount} has ${pageTemplates} templates. Total fetched so far: ${allTemplates.length}.`);

                // If we got fewer results than maxReturn, we've reached the end
                if (pageTemplates < maxReturn) {
                    moreResults = false;
                    console.log(`DEBUG (All Templates): Received fewer than ${maxReturn} results (${pageTemplates}), ending pagination.`);
                } else {
                    // Increment the offset for the next page
                    offset += maxReturn;
                }
            } else {
                console.warn("WARN (All Templates): Unexpected response format or API error");
                if (response.data?.errors) {
                    console.error("API Errors:", JSON.stringify(response.data.errors));
                }
                moreResults = false;
            }
        }

        console.log(`DEBUG (All Templates): Completed fetch. Returning ${allTemplates.length} templates as ${format || 'json'}.`);

        // If CSV format is requested, convert and return as CSV
        if (format === 'csv') {
            // Define fields for CSV (you can customize this based on what you want in the CSV)
            const fields = [
                'id',
                'name',
                'description',
                'subject',
                'fromName',
                'fromEmail',
                'replyTo',
                'status',
                'createdAt',
                'updatedAt',
                {
                    label: 'Folder ID',
                    value: row => row.folder?.id || ''
                },
                {
                    label: 'Folder Name',
                    value: row => row.folder?.value || ''
                },
                {
                    label: 'Folder Type',
                    value: row => row.folder?.type || ''
                }
            ];

            // Process timestamps and handle nested properties before conversion
            const processedTemplates = allTemplates.map(template => {
                // Create a copy of the template to avoid modifying the original
                const processedTemplate = { ...template };

                // Format dates nicely if they exist
                if (processedTemplate.createdAt) {
                    processedTemplate.createdAt = new Date(processedTemplate.createdAt).toLocaleString();
                }
                if (processedTemplate.updatedAt) {
                    processedTemplate.updatedAt = new Date(processedTemplate.updatedAt).toLocaleString();
                }

                return processedTemplate;
            });

            try {
                // Convert to CSV
                const json2csvParser = new Parser({ fields });
                const csv = json2csvParser.parse(processedTemplates);

                // Set headers for CSV download
                res.setHeader('Content-Type', 'text/csv');
                res.setHeader('Content-Disposition', 'attachment; filename=marketo_templates.csv');

                return res.send(csv);
            } catch (csvError) {
                console.error("ERROR (All Templates): Failed to convert to CSV:", csvError.message);
                return res.status(500).json({
                    success: false,
                    message: `Failed to convert to CSV: ${csvError.message}`
                });
            }
        }

        // Default: return as JSON
        return res.status(200).json({
            success: true,
            message: `Successfully fetched all ${allTemplates.length} templates.`,
            totalTemplates: allTemplates.length,
            templates: allTemplates
        });

    } catch (error) {
        console.error("ERROR (All Templates): Failed to fetch all templates:", error.message);
        return res.status(500).json({
            success: false,
            message: error.message || 'An unexpected error occurred while fetching all templates.'
        });
    }
});

// Create a simple exports view
// Create views/exports.ejs with:
/*
<div class="container mt-4">
    <h1><%= title %></h1>
    <p class="lead">View and download your exported Marketo templates.</p>
    <hr>

    <% if (exports.length === 0) { %>
        <div class="alert alert-info">No exports found.</div>
    <% } else { %>
        <div class="list-group">
            <% exports.forEach(function(export_item) { %>
                <div class="list-group-item">
                    <div class="d-flex w-100 justify-content-between">
                        <h5 class="mb-1"><%= export_item %></h5>
                    </div>
                    <% if (export_item.endsWith('.zip')) { %>
                        <a href="/marketo-export/download-archive/<%= export_item %>" class="btn btn-sm btn-primary mt-2">
                            Download ZIP
                        </a>
                    <% } else { %>
                        <a href="/marketo-export/create-archive/<%= export_item %>" class="btn btn-sm btn-primary mt-2">
                            Create ZIP & Download
                        </a>
                    <% } %>
                </div>
            <% }); %>
        </div>
    <% } %>

    <div class="mt-4">
        <a href="/marketo-export" class="btn btn-secondary">Back to Marketo Export</a>
    </div>
</div>
*/

router.get('/template-count', async (req, res) => {
    const restUrl = process.env.MARKETO_REST_URL;
    if (!restUrl) {
        console.error("ERROR: MARKETO_REST_URL is not configured.");
        return res.status(500).json({ success: false, message: "Marketo REST API URL is not configured." });
    }

    try {
        const accessToken = await getMarketoAccessToken(req);
        const templatesEndpoint = `${restUrl}/asset/v1/emails.json`;

        let totalTemplates = 0;
        let nextPageToken = null;
        const MAX_PAGES = 10; // Set a reasonable limit
        let pageCount = 0;
        let hasMore = false;

        console.log("DEBUG: Starting verification of total template count...");

        do {
            pageCount++;
            if (pageCount > MAX_PAGES) {
                console.warn(`WARN: Reached MAX_PAGES (${MAX_PAGES}). Stopping pagination for count check.`);
                break;
            }

            let params = { maxReturn: 200 };
            if (nextPageToken) {
                params.nextPageToken = nextPageToken;
                console.log(`DEBUG: Fetching count page ${pageCount} using token: ${nextPageToken}`);
            } else {
                console.log(`DEBUG: Fetching count page ${pageCount} (first page)...`);
            }

            const response = await axios.get(templatesEndpoint, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                params: params
            });

            if (response.data && Array.isArray(response.data.result)) {
                const pageTemplates = response.data.result.length;
                totalTemplates += pageTemplates;
                nextPageToken = response.data.nextPageToken;
                hasMore = response.data.moreResult === true;

                console.log(`DEBUG: Page ${pageCount} has ${pageTemplates} templates. Total so far: ${totalTemplates}.`);
                console.log(`DEBUG: More results? ${hasMore}. Next token? ${nextPageToken ? 'Yes' : 'No'}`);
            } else {
                console.warn("WARN: Unexpected response format:", JSON.stringify(response.data));
                break;
            }

        } while (hasMore && nextPageToken);

        return res.status(200).json({
            success: true,
            message: `Found ${totalTemplates} templates across ${pageCount} pages.`,
            totalTemplates: totalTemplates,
            pagesChecked: pageCount,
            hasMore: hasMore && nextPageToken !== null
        });

    } catch (error) {
        console.error("ERROR: Failed to count templates:", error.message);
        return res.status(500).json({
            success: false,
            message: error.message || 'An unexpected error occurred while counting templates.'
        });
    }
});
// Route to create an archive for an existing export directory
router.get('/create-archive/:dirname', async (req, res) => {
    try {
        const dirName = req.params.dirname;
        const exportDir = path.join(process.cwd(), 'exports', dirName);

        if (!fs.existsSync(exportDir) || !fs.statSync(exportDir).isDirectory()) {
            return res.status(404).send('Export directory not found');
        }

        const zipPath = await createZipArchive(exportDir);
        res.redirect(`/marketo-export/download-archive/${path.basename(zipPath)}`);
    } catch (err) {
        console.error("Error creating archive:", err);
        res.status(500).send("Error creating archive: " + err.message);
    }
});

// Route to download the ZIP archive
router.get('/download-archive/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(process.cwd(), 'exports', filename);

    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('Archive not found');
    }
});

router.get('/email/:id/screenshot', async (req, res) => {
    const emailId = req.params.id;
    let browser = null;

    try {
        // Use new API method to get HTML content
        const rawHtmlContent = await getCompleteEmailContent(req, emailId);

        if (!rawHtmlContent) {
            return res.status(404).send("No HTML content found");
        }

        const wrappedHtml = enhancedMarketoEmailHtml(rawHtmlContent, emailId);

        browser = await puppeteer.launch({
            args: ['--no-sandbox', '--disable-setuid-sandbox'],
            headless: true
        });

        const page = await browser.newPage();

        await page.setViewport({
            width: 650,
            height: 800,
            deviceScaleFactor: 1,
        });

        await page.setContent(wrappedHtml, {
            waitUntil: 'networkidle0',
            timeout: 30000
        });

        const screenshot = await page.screenshot({
            fullPage: true,
            type: 'png'
        });

        await browser.close();
        browser = null;

        res.setHeader('Content-Type', 'image/png');
        res.send(screenshot);
    } catch (error) {
        if (browser) await browser.close();
        res.status(500).send(`Error: ${error.message}`);
    }
});

module.exports = router;