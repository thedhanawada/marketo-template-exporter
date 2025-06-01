#!/usr/bin/env node

const { program } = require('commander');
const dotenv = require('dotenv');
const ora = require('ora');
const chalk = require('chalk');
const path = require('path');
const MarketoClient = require('./lib/marketo-client');

dotenv.config();

program
    .name('marketo-export')
    .description('Export Marketo email templates')
    .version('1.0.0');

program
    .command('export')
    .description('Export all Marketo email templates')
    .option('-o, --output <directory>', 'Output directory', './marketo-exports')
    .option('-z, --zip', 'Create ZIP archive of exports')
    .action(async (options) => {
        const spinner = ora('Starting export process').start();

        try {
            // Validate environment variables
            const requiredEnvVars = [
                'MARKETO_CLIENT_ID',
                'MARKETO_CLIENT_SECRET',
                'MARKETO_IDENTITY_URL',
                'MARKETO_REST_URL'
            ];

            const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);
            if (missingVars.length > 0) {
                throw new Error(`Missing required environment variables: ${missingVars.join(', ')}`);
            }

            const client = new MarketoClient({
                clientId: process.env.MARKETO_CLIENT_ID,
                clientSecret: process.env.MARKETO_CLIENT_SECRET,
                identityUrl: process.env.MARKETO_IDENTITY_URL,
                restUrl: process.env.MARKETO_REST_URL
            });

            spinner.text = 'Authenticating with Marketo';
            await client.authenticate();

            spinner.text = 'Fetching templates';
            const results = await client.exportAllTemplates(
                options.output,
                (progress) => {
                    if (typeof progress === 'number') {
                        spinner.text = `Fetched ${progress} templates`;
                    } else {
                        spinner.text = `Processed ${progress.successful + progress.failed}/${progress.total} templates`;
                    }
                }
            );

            if (options.zip) {
                spinner.text = 'Creating ZIP archive';
                const zipPath = path.join(options.output, `marketo-templates-${Date.now()}.zip`);
                await client.createArchive(options.output, zipPath);
            }

            spinner.succeed(chalk.green('Export completed successfully!'));
            console.log(chalk.cyan('\nSummary:'));
            console.log(chalk.white(`Total templates: ${results.total}`));
            console.log(chalk.green(`Successfully exported: ${results.successful}`));
            console.log(chalk.red(`Failed: ${results.failed}`));

            if (results.errors.length > 0) {
                console.log(chalk.yellow('\nErrors:'));
                results.errors.forEach(error => {
                    console.log(chalk.red(`  Template ${error.templateId}: ${error.error}`));
                });
            }

        } catch (error) {
            spinner.fail(chalk.red(`Export failed: ${error.message}`));
            process.exit(1);
        }
    });

program.parse(); 