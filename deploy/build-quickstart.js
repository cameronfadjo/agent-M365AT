const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
        ShadingType, PageNumber, PageBreak, ExternalHyperlink, LevelFormat } = require('docx');
const fs = require('fs');

const BLUE = "1B3A5C";
const LIGHT_BLUE = "D5E8F0";
const ACCENT = "2E75B6";
const GRAY = "666666";
const LIGHT_GRAY = "F5F5F5";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function heading(text, level) {
    return new Paragraph({
        heading: level,
        children: [new TextRun(text)]
    });
}

function para(text, opts = {}) {
    return new Paragraph({
        spacing: { after: 120 },
        ...opts,
        children: [new TextRun({ font: "Arial", size: 22, ...opts.run, text })]
    });
}

function boldPara(text, opts = {}) {
    return new Paragraph({
        spacing: { after: 120 },
        ...opts,
        children: [new TextRun({ font: "Arial", size: 22, bold: true, ...opts.run, text })]
    });
}

function stepNumber(num, title, description) {
    return [
        new Paragraph({
            spacing: { before: 300, after: 100 },
            children: [
                new TextRun({ font: "Arial", size: 36, bold: true, color: ACCENT, text: `Step ${num}: ` }),
                new TextRun({ font: "Arial", size: 28, bold: true, text: title }),
            ]
        }),
        para(description),
    ];
}

function infoBox(children) {
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1, color: ACCENT },
                            bottom: { style: BorderStyle.SINGLE, size: 1, color: ACCENT },
                            left: { style: BorderStyle.SINGLE, size: 6, color: ACCENT },
                            right: { style: BorderStyle.SINGLE, size: 1, color: ACCENT },
                        },
                        shading: { fill: "EBF5FB", type: ShadingType.CLEAR },
                        margins: { top: 120, bottom: 120, left: 200, right: 200 },
                        width: { size: 9360, type: WidthType.DXA },
                        children: children,
                    })
                ]
            })
        ]
    });
}

const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 22 } } },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 36, bold: true, font: "Arial", color: BLUE },
                paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: "Arial", color: ACCENT },
                paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Arial" },
                paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
        ]
    },
    numbering: {
        config: [
            {
                reference: "bullets",
                levels: [{
                    level: 0, format: LevelFormat.BULLET, text: "\u2022",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "numbers",
                levels: [{
                    level: 0, format: LevelFormat.DECIMAL, text: "%1.",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
        ]
    },
    sections: [
        // ============================================================
        // COVER / TITLE PAGE
        // ============================================================
        {
            properties: {
                page: {
                    size: { width: 12240, height: 15840 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                }
            },
            children: [
                new Paragraph({ spacing: { before: 2400 } }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                    children: [new TextRun({ font: "Arial", size: 56, bold: true, color: BLUE, text: "Refresh Agent" })]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 100 },
                    children: [new TextRun({ font: "Arial", size: 28, color: ACCENT, text: "Customer Deployment Quick Start Guide" })]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
                    spacing: { after: 400 },
                    children: []
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 100 },
                    children: [new TextRun({ font: "Arial", size: 22, color: GRAY, text: "AI-Powered Document Research & Generation for K-12 School Districts" })]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 600 },
                    children: [new TextRun({ font: "Arial", size: 20, color: GRAY, text: "Version 6.0  |  February 2026" })]
                }),
                infoBox([
                    new Paragraph({
                        spacing: { after: 80 },
                        children: [new TextRun({ font: "Arial", size: 22, bold: true, text: "What is Refresh Agent?" })]
                    }),
                    para("Refresh Agent helps educators create updated versions of recurring documents " +
                         "(back-to-school letters, budget memos, policy updates) by analyzing previous " +
                         "versions in OneDrive, discovering recent organizational changes, and generating " +
                         "new versions grounded in document history and context."),
                    new Paragraph({
                        spacing: { after: 80 },
                        children: [new TextRun({ font: "Arial", size: 22, bold: true, text: "Deployment time: ~10 minutes" })]
                    }),
                ]),
                new Paragraph({ children: [new PageBreak()] }),

                // ============================================================
                // WHAT YOU NEED
                // ============================================================
                heading("Before You Begin", HeadingLevel.HEADING_1),
                para("Your CDW implementation team will provide the deployment package and run the automated setup. " +
                     "The following items are needed from your district:"),
                new Paragraph({ spacing: { after: 200 } }),
                new Table({
                    width: { size: 9360, type: WidthType.DXA },
                    columnWidths: [4000, 5360],
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 4000, type: WidthType.DXA },
                                    shading: { fill: BLUE, type: ShadingType.CLEAR },
                                    margins: cellMargins,
                                    children: [new Paragraph({ children: [new TextRun({ bold: true, color: "FFFFFF", font: "Arial", size: 22, text: "Requirement" })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 5360, type: WidthType.DXA },
                                    shading: { fill: BLUE, type: ShadingType.CLEAR },
                                    margins: cellMargins,
                                    children: [new Paragraph({ children: [new TextRun({ bold: true, color: "FFFFFF", font: "Arial", size: 22, text: "Details" })] })]
                                }),
                            ]
                        }),
                        ...([
                            ["Microsoft 365 tenant", "With M365 Copilot licenses for agent users"],
                            ["Azure subscription", "With permission to create resources (Contributor role)"],
                            ["Azure OpenAI resource", "With a gpt-4o-mini (or similar) model deployed"],
                            ["Azure Storage account", "With a blob container for generated documents"],
                            ["Global Admin access", "To grant admin consent for Graph API permissions"],
                            ["Teams Admin access", "To upload and assign the agent app package"],
                        ].map(([req, detail], i) =>
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders, width: { size: 4000, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [new Paragraph({ children: [new TextRun({ bold: true, font: "Arial", size: 22, text: req })] })]
                                    }),
                                    new TableCell({
                                        borders, width: { size: 5360, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [para(detail)]
                                    }),
                                ]
                            })
                        ))
                    ]
                }),

                new Paragraph({ children: [new PageBreak()] }),

                // ============================================================
                // DEPLOYMENT STEPS
                // ============================================================
                heading("Deployment Steps", HeadingLevel.HEADING_1),

                para("CDW handles the heavy lifting. Here is what happens during deployment and what you need to do."),

                // STEP 1
                ...stepNumber("1", "Automated Infrastructure Setup", "CDW runs the deployment script against your Azure subscription. This creates:"),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Azure Function App (hosts the Refresh Agent backend)" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Entra ID app registration (OAuth 2.0 for secure OneDrive access)" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "All security scopes and authorized client applications" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Python code deployment with all 13 API endpoints" })] }),

                infoBox([
                    para("Your action: Provide your Azure subscription ID, Azure OpenAI credentials, and storage connection string to your CDW implementation team."),
                ]),

                // STEP 2
                ...stepNumber("2", "Grant Admin Consent", "After the automated setup completes, a Global Admin must grant consent for the Graph API permissions the agent needs."),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Go to Azure Portal \u2192 Microsoft Entra ID \u2192 App registrations" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Find the app named \"Refresh Agent - [your-app-name]\"" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Go to API permissions" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Click \"Grant admin consent\" \u2192 Confirm" })] }),

                para("The agent requests two permissions:"),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Files.ReadWrite \u2014 to read and write documents in users\u2019 OneDrive" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "User.Read \u2014 to read the user\u2019s basic profile" })] }),

                // STEP 3
                ...stepNumber("3", "Install the Agent", "CDW provides a .zip app package. A Teams Admin uploads it to your tenant."),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Go to Teams Admin Center (admin.teams.microsoft.com)" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Navigate to Teams apps \u2192 Manage apps" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Click \"Upload new app\" \u2192 select the .zip file" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Assign the app to users or groups who should have access" })] }),

                infoBox([
                    para("That\u2019s it. The agent is now available in M365 Copilot for assigned users."),
                ]),

                new Paragraph({ children: [new PageBreak()] }),

                // ============================================================
                // USING THE AGENT
                // ============================================================
                heading("Using the Refresh Agent", HeadingLevel.HEADING_1),

                para("Once installed, users access the Refresh Agent through M365 Copilot in Teams or Outlook."),

                heading("Getting Started", HeadingLevel.HEADING_2),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Open M365 Copilot in Teams (or Outlook)" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Find and select the Refresh Agent" })] }),
                new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Ask for a document: \"I need this year\u2019s back-to-school letter\"" })] }),

                heading("What Happens Behind the Scenes", HeadingLevel.HEADING_2),
                para("The agent follows a 6-step workflow automatically:"),

                new Table({
                    width: { size: 9360, type: WidthType.DXA },
                    columnWidths: [1200, 3000, 5160],
                    rows: [
                        new TableRow({
                            children: ["Step", "Action", "What It Does"].map((h, i) =>
                                new TableCell({
                                    borders, width: { size: [1200, 3000, 5160][i], type: WidthType.DXA },
                                    shading: { fill: BLUE, type: ShadingType.CLEAR },
                                    margins: cellMargins,
                                    children: [new Paragraph({ children: [new TextRun({ bold: true, color: "FFFFFF", font: "Arial", size: 20, text: h })] })]
                                })
                            )
                        }),
                        ...([
                            ["1", "Understand", "Parses your request to identify document type and search keywords"],
                            ["2", "Search", "Finds previous versions of the document in your OneDrive"],
                            ["3", "Context", "Searches for recent organizational documents (memos, announcements)"],
                            ["4", "Analyze", "Compares versions to find patterns, changes, and new elements"],
                            ["5", "Generate", "Creates an updated version incorporating all findings"],
                            ["6", "Save", "Saves the new document to your OneDrive"],
                        ].map(([step, action, desc], i) =>
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders, width: { size: 1200, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ bold: true, font: "Arial", size: 20, color: ACCENT, text: step })] })]
                                    }),
                                    new TableCell({
                                        borders, width: { size: 3000, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [new Paragraph({ children: [new TextRun({ bold: true, font: "Arial", size: 20, text: action })] })]
                                    }),
                                    new TableCell({
                                        borders, width: { size: 5160, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [para(desc)]
                                    }),
                                ]
                            })
                        ))
                    ]
                }),

                heading("Example Prompts", HeadingLevel.HEADING_2),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "\"I need this year\u2019s back-to-school letter\"" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "\"Update our budget memo for the new fiscal year\"" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "\"Generate an updated student AUP acknowledgment form\"" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "\"Help me refresh our staff technology notice\"" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "\"Create a new version of a recurring document\"" })] }),

                new Paragraph({ children: [new PageBreak()] }),

                // ============================================================
                // SECURITY & DATA
                // ============================================================
                heading("Security & Data Privacy", HeadingLevel.HEADING_1),

                para("The Refresh Agent is designed for enterprise security and data isolation."),

                heading("Data Stays in Your Tenant", HeadingLevel.HEADING_2),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "All infrastructure runs in your Azure subscription" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Documents are read from and saved to the user\u2019s own OneDrive" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Generated documents are stored in your Azure Blob Storage" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "No data is shared with CDW or any third party" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "AI processing uses your Azure OpenAI resource (within your subscription)" })] }),

                heading("Authentication", HeadingLevel.HEADING_2),
                para("The agent uses OAuth 2.0 On-Behalf-Of (OBO) flow. When a user interacts with the agent, " +
                     "their existing M365 sign-in is used to securely access their OneDrive. The agent can only " +
                     "access files the user already has permission to see. No additional passwords or logins are required."),

                heading("Permissions", HeadingLevel.HEADING_2),
                para("The agent requests the minimum permissions needed:"),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "Files.ReadWrite \u2014 Read previous document versions, save generated documents" })] }),
                new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 120 },
                    children: [new TextRun({ font: "Arial", size: 22, text: "User.Read \u2014 Identify the user for OneDrive access" })] }),

                new Paragraph({ children: [new PageBreak()] }),

                // ============================================================
                // SUPPORT
                // ============================================================
                heading("Support & Troubleshooting", HeadingLevel.HEADING_1),

                heading("Common Issues", HeadingLevel.HEADING_2),
                new Table({
                    width: { size: 9360, type: WidthType.DXA },
                    columnWidths: [3500, 5860],
                    rows: [
                        new TableRow({
                            children: ["Issue", "Resolution"].map((h, i) =>
                                new TableCell({
                                    borders, width: { size: [3500, 5860][i], type: WidthType.DXA },
                                    shading: { fill: BLUE, type: ShadingType.CLEAR },
                                    margins: cellMargins,
                                    children: [new Paragraph({ children: [new TextRun({ bold: true, color: "FFFFFF", font: "Arial", size: 20, text: h })] })]
                                })
                            )
                        }),
                        ...([
                            ["Agent not appearing in Copilot", "Ensure the app is assigned to the user in Teams Admin Center"],
                            ["\"Could not connect\" error", "Verify admin consent was granted for API permissions"],
                            ["OneDrive search returns no results", "Try simpler search terms; the agent works best with 2-3 keyword queries"],
                            ["Document generation fails", "Check that the Azure OpenAI resource is accessible and the API key is valid"],
                            ["Cannot save to OneDrive", "Verify the Files.ReadWrite permission has admin consent"],
                        ].map(([issue, fix], i) =>
                            new TableRow({
                                children: [
                                    new TableCell({
                                        borders, width: { size: 3500, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [new Paragraph({ children: [new TextRun({ font: "Arial", size: 20, text: issue })] })]
                                    }),
                                    new TableCell({
                                        borders, width: { size: 5860, type: WidthType.DXA },
                                        shading: { fill: i % 2 === 0 ? LIGHT_GRAY : "FFFFFF", type: ShadingType.CLEAR },
                                        margins: cellMargins,
                                        children: [para(fix)]
                                    }),
                                ]
                            })
                        ))
                    ]
                }),

                new Paragraph({ spacing: { before: 300 } }),
                heading("Contact CDW Support", HeadingLevel.HEADING_2),
                para("For deployment assistance or technical issues, contact your CDW implementation team."),
            ]
        }
    ]
});

Packer.toBuffer(doc).then(buffer => {
    const outPath = "/sessions/awesome-admiring-noether/mnt/PUFSD - TBO/Refresh-M365AT/deploy/Refresh-Agent-QuickStart.docx";
    fs.writeFileSync(outPath, buffer);
    console.log("Quick Start Guide created: " + outPath);
});
