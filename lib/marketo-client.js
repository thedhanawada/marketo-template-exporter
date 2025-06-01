const axios = require('axios');
const fs = require('fs').promises;
const path = require('path');
const archiver = require('archiver');

class MarketoClient {
    constructor(config) {
        this.clientId = config.clientId;
        this.clientSecret = config.clientSecret;
        this.identityUrl = config.identityUrl;
        this.restUrl = config.restUrl;
        this.accessToken = null;
        this.tokenExpiry = null;
    }

    async authenticate() {
        const tokenRequestUrl = `${this.identityUrl}?grant_type=client_credentials&client_id=${this.clientId}&client_secret=${this.clientSecret}`;
        
        try {
            const response = await axios.get(tokenRequestUrl);
            if (response.data && response.data.access_token) {
                this.accessToken = response.data.access_token;
                this.tokenExpiry = Date.now() + (response.data.expires_in * 1000);
                return this.accessToken;
            }
            throw new Error('Invalid authentication response');
        } catch (error) {
            throw new Error(`Authentication failed: ${error.message}`);
        }
    }

    async getAccessToken() {
        if (!this.accessToken || Date.now() >= this.tokenExpiry - 5 * 60 * 1000) {
            await this.authenticate();
        }
        return this.accessToken;
    }

    async getAllTemplates(progressCallback) {
        const templates = [];
        let offset = 0;
        const maxReturn = 200;

        while (true) {
            const accessToken = await this.getAccessToken();
            const response = await axios.get(`${this.restUrl}/asset/v1/emails.json`, {
                headers: { 'Authorization': `Bearer ${accessToken}` },
                params: { maxReturn, offset }
            });

            if (!response.data.success) {
                throw new Error('Failed to fetch templates');
            }

            templates.push(...response.data.result);
            
            if (progressCallback) {
                progressCallback(templates.length);
            }

            if (response.data.result.length < maxReturn) {
                break;
            }
            offset += maxReturn;
        }

        return templates;
    }

    async getTemplateContent(templateId) {
        const accessToken = await this.getAccessToken();
        const response = await axios.get(
            `${this.restUrl}/asset/v1/email/${templateId}/content.json`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );

        if (!response.data.success) {
            throw new Error(`Failed to fetch template content for ID ${templateId}`);
        }

        return this.extractHtmlContent(response.data.result);
    }

    extractHtmlContent(contentSections) {
        let htmlContent = '';
        for (const section of contentSections) {
            if (section.value && Array.isArray(section.value)) {
                for (const item of section.value) {
                    if (item.type === 'HTML' && item.value) {
                        htmlContent += item.value + '\n';
                    }
                }
            }
            if (section.contentType === 'html' && section.content) {
                htmlContent += section.content + '\n';
            }
        }
        return htmlContent;
    }

    async exportTemplate(template, outputDir) {
        const templateDir = path.join(outputDir, `template_${template.id}`);
        await fs.mkdir(templateDir, { recursive: true });

        // Save metadata
        const metadata = {
            id: template.id,
            name: template.name,
            status: template.status,
            createdAt: template.createdAt,
            updatedAt: template.updatedAt,
            folder: template.folder,
            subject: template.subject,
            fromName: template.fromName,
            fromEmail: template.fromEmail
        };

        await fs.writeFile(
            path.join(templateDir, 'metadata.json'),
            JSON.stringify(metadata, null, 2)
        );

        // Get and save HTML content
        try {
            const htmlContent = await this.getTemplateContent(template.id);
            if (htmlContent) {
                await fs.writeFile(
                    path.join(templateDir, `${template.id}.html`),
                    htmlContent
                );
            }
        } catch (error) {
            await fs.writeFile(
                path.join(templateDir, 'error.txt'),
                `Error retrieving HTML content: ${error.message}`
            );
        }

        return templateDir;
    }

    async exportAllTemplates(outputDir, progressCallback) {
        const templates = await this.getAllTemplates(progressCallback);
        const results = {
            total: templates.length,
            successful: 0,
            failed: 0,
            errors: []
        };

        for (const template of templates) {
            try {
                await this.exportTemplate(template, outputDir);
                results.successful++;
            } catch (error) {
                results.failed++;
                results.errors.push({
                    templateId: template.id,
                    error: error.message
                });
            }
            if (progressCallback) {
                progressCallback(results);
            }
        }

        return results;
    }

    async createArchive(sourceDir, outputPath) {
        return new Promise((resolve, reject) => {
            const output = fs.createWriteStream(outputPath);
            const archive = archiver('zip', { zlib: { level: 9 } });

            output.on('close', () => resolve(outputPath));
            archive.on('error', reject);

            archive.pipe(output);
            archive.directory(sourceDir, false);
            archive.finalize();
        });
    }
}

module.exports = MarketoClient; 