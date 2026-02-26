/**
 * CDW Agent — Deployment Portal Backend
 *
 * Node.js/Express server that:
 *   1. Serves the portal UI (index.html)
 *   2. Executes Deploy-Full.ps1 via SSE endpoint
 *   3. Streams real-time progress back to the browser
 *   4. Serves the built agent package (.zip) for download
 *
 * Usage:
 *   node server.js                    # Start on default port 3000
 *   PORT=8080 node server.js          # Start on custom port
 *
 * Prerequisites:
 *   - Node.js 18+
 *   - PowerShell 7+ (pwsh) installed and on PATH
 *   - Azure CLI (az) installed and on PATH
 *   - Azure Functions Core Tools (func) installed and on PATH
 *   - Microsoft.Graph PowerShell module installed
 */

const express = require('express');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Resolve paths
const PORTAL_DIR = __dirname;
const DEPLOY_DIR = path.resolve(PORTAL_DIR, '..');
const PROJECT_ROOT = path.resolve(DEPLOY_DIR, '..');

// ============================================================================
// STATIC FILES
// ============================================================================

app.use(express.static(PORTAL_DIR));

// ============================================================================
// SSE DEPLOYMENT ENDPOINT
// ============================================================================

app.get('/api/deploy', (req, res) => {
    // SSE headers
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'X-Accel-Buffering': 'no',
    });

    // Flush headers immediately
    res.flushHeaders();

    // Helper: send SSE event
    function sendEvent(eventType, data) {
        res.write(`event: ${eventType}\ndata: ${JSON.stringify(data)}\n\n`);
    }

    // Extract parameters from query string
    const {
        tenantId,
        subscriptionId,
        appName,
        region,
        azureOpenAIEndpoint,
        azureOpenAIKey,
        azureOpenAIDeployment,
        azureOpenAIDeploymentLarge,
        azureOpenAIApiVersion,
        storageConnectionString,
        storageContainerName,
        sasTokenExpiryHours,
    } = req.query;

    // Validate required fields
    const missing = [];
    if (!appName) missing.push('appName');
    if (!region) missing.push('region');
    if (!azureOpenAIEndpoint) missing.push('azureOpenAIEndpoint');
    if (!azureOpenAIKey) missing.push('azureOpenAIKey');
    if (!storageConnectionString) missing.push('storageConnectionString');

    if (missing.length > 0) {
        sendEvent('error_event', {
            id: 'auth',
            message: `Missing required fields: ${missing.join(', ')}`,
        });
        res.end();
        return;
    }

    // Build PowerShell arguments
    const scriptPath = path.join(DEPLOY_DIR, 'Deploy-Full.ps1');

    if (!fs.existsSync(scriptPath)) {
        sendEvent('error_event', {
            id: 'auth',
            message: `Deploy-Full.ps1 not found at: ${scriptPath}`,
        });
        res.end();
        return;
    }

    const psArgs = [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy', 'Bypass',
        '-File', scriptPath,
        '-AppName', appName,
        '-Region', region,
        '-AzureOpenAIEndpoint', azureOpenAIEndpoint,
        '-AzureOpenAIKey', azureOpenAIKey,
        '-StorageConnectionString', storageConnectionString,
    ];

    // Optional parameters
    if (azureOpenAIDeployment) {
        psArgs.push('-AzureOpenAIDeployment', azureOpenAIDeployment);
    }
    if (azureOpenAIDeploymentLarge) {
        psArgs.push('-AzureOpenAIDeploymentLarge', azureOpenAIDeploymentLarge);
    }
    if (azureOpenAIApiVersion) {
        psArgs.push('-AzureOpenAIApiVersion', azureOpenAIApiVersion);
    }
    if (storageContainerName) {
        psArgs.push('-StorageContainerName', storageContainerName);
    }
    if (sasTokenExpiryHours) {
        psArgs.push('-SasTokenExpiryHours', sasTokenExpiryHours);
    }

    // Determine PowerShell executable
    const pwsh = process.platform === 'win32' ? 'pwsh.exe' : 'pwsh';

    console.log(`[deploy] Starting deployment for: ${appName}`);
    console.log(`[deploy] Script: ${scriptPath}`);

    // Track which deployment step we're on based on output parsing
    let currentStepId = 'auth';
    let deploymentResult = {};

    // Spawn the PowerShell process
    const child = spawn(pwsh, psArgs, {
        cwd: DEPLOY_DIR,
        env: {
            ...process.env,
            // Ensure output encoding works
            PYTHONIOENCODING: 'utf-8',
        },
    });

    // ── Parse output lines and map to SSE events ──

    function processLine(line) {
        const trimmed = line.trim();
        if (!trimmed) return;

        // ── Step transitions ──
        // [1/5] Authenticating to Azure...
        if (trimmed.includes('[1/5]')) {
            currentStepId = 'auth';
            sendEvent('step', { id: 'auth', status: 'active', message: 'Authenticating to Azure...' });
        }
        // [2/5] Deploying Azure infrastructure...
        else if (trimmed.includes('[2/5]')) {
            sendEvent('step', { id: 'auth', status: 'completed', message: 'Authenticated' });
            currentStepId = 'infra';
            sendEvent('step', { id: 'infra', status: 'active', message: 'Deploying Azure infrastructure...' });
        }
        // [3/5] Configuring Entra ID...
        else if (trimmed.includes('[3/5]')) {
            sendEvent('step', { id: 'infra', status: 'completed', message: 'Infrastructure deployed' });
            currentStepId = 'entra';
            sendEvent('step', { id: 'entra', status: 'active', message: 'Configuring Entra ID...' });
        }
        // [4/5] Deploying Python code...
        else if (trimmed.includes('[4/5]')) {
            sendEvent('step', { id: 'entra', status: 'completed', message: 'Entra ID configured' });
            currentStepId = 'code';
            sendEvent('step', { id: 'code', status: 'active', message: 'Deploying Python code...' });
        }
        // [5/5] Building agent package...
        else if (trimmed.includes('[5/5]')) {
            sendEvent('step', { id: 'code', status: 'completed', message: 'Code deployed' });
            currentStepId = 'package';
            sendEvent('step', { id: 'package', status: 'active', message: 'Building agent package...' });
        }
        // DEPLOYMENT COMPLETE
        else if (trimmed.includes('DEPLOYMENT COMPLETE')) {
            sendEvent('step', { id: 'package', status: 'completed', message: 'Package built' });
        }
        // Prerequisites check
        else if (trimmed.includes('[0/5]')) {
            sendEvent('step', { id: 'auth', status: 'active', message: 'Checking prerequisites...' });
        }

        // ── Capture key values from output ──
        if (trimmed.includes('Subscription:')) {
            const match = trimmed.match(/Subscription:\s*(.+)\s*\((.+)\)/);
            if (match) {
                deploymentResult.subscriptionName = match[1].trim();
                deploymentResult.subscriptionId = match[2].trim();
            }
        }
        if (trimmed.includes('Function App:') && trimmed.includes('https://')) {
            const match = trimmed.match(/https:\/\/[^\s]+/);
            if (match) deploymentResult.functionUrl = match[0];
        }
        if (trimmed.includes('Client ID:') && !trimmed.includes('PLACEHOLDER')) {
            const match = trimmed.match(/Client ID:\s*([a-f0-9-]+)/i);
            if (match) deploymentResult.clientId = match[1];
        }
        if (trimmed.includes('Tenant ID:') && !trimmed.includes('PLACEHOLDER')) {
            const match = trimmed.match(/Tenant ID:\s*([a-f0-9-]+)/i);
            if (match) deploymentResult.tenantId = match[1];
        }
        if (trimmed.includes('Package:') && trimmed.includes('.zip')) {
            const match = trimmed.match(/Package:\s*(.+\.zip)/);
            if (match) deploymentResult.packagePath = match[1].trim();
        }

        // ── Always send as log ──
        sendEvent('log', { id: currentStepId, message: trimmed });
    }

    // Buffer partial lines from stdout/stderr
    let stdoutBuffer = '';
    let stderrBuffer = '';

    child.stdout.on('data', (data) => {
        stdoutBuffer += data.toString();
        const lines = stdoutBuffer.split('\n');
        stdoutBuffer = lines.pop(); // Keep the incomplete last line
        lines.forEach(processLine);
    });

    child.stderr.on('data', (data) => {
        stderrBuffer += data.toString();
        const lines = stderrBuffer.split('\n');
        stderrBuffer = lines.pop();
        lines.forEach((line) => {
            const trimmed = line.trim();
            if (!trimmed) return;
            // PowerShell Write-Information goes to stderr in some modes
            // Only treat as error if it looks like one
            if (trimmed.match(/^(ERROR|FATAL|Exception|throw|Unhandled)/i)) {
                sendEvent('log', { id: currentStepId, message: `ERROR: ${trimmed}` });
            } else {
                processLine(trimmed);
            }
        });
    });

    child.on('close', (code) => {
        // Flush remaining buffers
        if (stdoutBuffer.trim()) processLine(stdoutBuffer);
        if (stderrBuffer.trim()) processLine(stderrBuffer);

        console.log(`[deploy] Process exited with code: ${code}`);

        if (code === 0) {
            // Try to load deployment-record.json for complete results
            const recordPath = path.join(DEPLOY_DIR, 'deployment-record.json');
            if (fs.existsSync(recordPath)) {
                try {
                    const record = JSON.parse(fs.readFileSync(recordPath, 'utf-8'));
                    deploymentResult = { ...deploymentResult, ...record };
                } catch (e) {
                    console.warn('[deploy] Could not parse deployment-record.json:', e.message);
                }
            }

            // Fill in defaults
            deploymentResult.functionUrl = deploymentResult.functionUrl || deploymentResult.FunctionUrl || `https://${appName}.azurewebsites.net`;
            deploymentResult.resourceGroup = deploymentResult.resourceGroup || `rg-${appName}`;
            deploymentResult.tenantId = deploymentResult.tenantId || deploymentResult.TenantId || tenantId;
            deploymentResult.clientId = deploymentResult.clientId || deploymentResult.ClientId || '';
            deploymentResult.packagePath = deploymentResult.packagePath || deploymentResult.PackagePath || `cdw-agent-${appName}.zip`;

            sendEvent('complete', {
                functionUrl: deploymentResult.functionUrl,
                resourceGroup: deploymentResult.resourceGroup,
                clientId: deploymentResult.clientId,
                tenantId: deploymentResult.tenantId,
                packagePath: deploymentResult.packagePath,
            });
        } else {
            sendEvent('error_event', {
                id: currentStepId,
                message: `Deployment failed at step "${currentStepId}" (exit code ${code}). Check the logs above for details.`,
            });
        }

        res.end();
    });

    child.on('error', (err) => {
        console.error(`[deploy] Failed to start process:`, err);
        sendEvent('error_event', {
            id: 'auth',
            message: `Failed to start PowerShell: ${err.message}. Is 'pwsh' installed and on PATH?`,
        });
        res.end();
    });

    // Handle client disconnect
    req.on('close', () => {
        console.log(`[deploy] Client disconnected, killing process...`);
        child.kill('SIGTERM');
    });
});

// ============================================================================
// PACKAGE DOWNLOAD ENDPOINT
// ============================================================================

app.get('/api/download-package', (req, res) => {
    const { appName } = req.query;

    if (!appName) {
        return res.status(400).json({ error: 'appName is required' });
    }

    // Check multiple possible locations
    const candidates = [
        path.join(DEPLOY_DIR, `cdw-agent-${appName}.zip`),
        path.join(DEPLOY_DIR, 'cdw-agent-package.zip'),
    ];

    const packagePath = candidates.find(p => fs.existsSync(p));

    if (!packagePath) {
        return res.status(404).json({
            error: 'Package not found',
            searched: candidates,
        });
    }

    res.download(packagePath, path.basename(packagePath));
});

// ============================================================================
// HEALTH / STATUS
// ============================================================================

app.get('/api/status', (req, res) => {
    // Check if prerequisites are available
    const { execSync } = require('child_process');
    const checks = {};

    try {
        execSync('pwsh --version', { timeout: 5000 });
        checks.pwsh = true;
    } catch {
        checks.pwsh = false;
    }

    try {
        execSync('az version', { timeout: 5000 });
        checks.az = true;
    } catch {
        checks.az = false;
    }

    try {
        execSync('func --version', { timeout: 5000 });
        checks.func = true;
    } catch {
        checks.func = false;
    }

    checks.deployScript = fs.existsSync(path.join(DEPLOY_DIR, 'Deploy-Full.ps1'));
    checks.projectRoot = fs.existsSync(path.join(PROJECT_ROOT, 'azure_function'));

    res.json({
        status: Object.values(checks).every(v => v) ? 'ready' : 'missing_prerequisites',
        checks,
        paths: {
            deployDir: DEPLOY_DIR,
            projectRoot: PROJECT_ROOT,
        },
    });
});

// ============================================================================
// FALLBACK — serve index.html for any unknown route
// ============================================================================

app.get('*', (req, res) => {
    res.sendFile(path.join(PORTAL_DIR, 'index.html'));
});

// ============================================================================
// START
// ============================================================================

app.listen(PORT, () => {
    console.log('');
    console.log('  ╔══════════════════════════════════════════════════════════╗');
    console.log('  ║       CDW Agent — Deployment Portal                 ║');
    console.log('  ╠══════════════════════════════════════════════════════════╣');
    console.log(`  ║  URL:  http://localhost:${PORT}                            ║`);
    console.log('  ║                                                          ║');
    console.log('  ║  Endpoints:                                              ║');
    console.log('  ║    GET  /                    Portal UI                   ║');
    console.log('  ║    GET  /api/deploy          SSE deployment stream       ║');
    console.log('  ║    GET  /api/download-package Download .zip package     ║');
    console.log('  ║    GET  /api/status          Prerequisites check         ║');
    console.log('  ╚══════════════════════════════════════════════════════════╝');
    console.log('');
});
