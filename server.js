'use strict';

const express = require('express');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const {
    Client,
    LocalAuth,
    MessageAck
} = require('whatsapp-web.js');
const QRCode = require('qrcode');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

/*
 * ---------------------------------------------------------
 * WhatsApp client
 * ---------------------------------------------------------
 */

let whatsappState = 'starting';
let whatsappQr = null;
let ack_messages = [];

const client = new Client({
    authStrategy: new LocalAuth({
        dataPath: path.join(__dirname, '.wwebjs_auth')
    }),
    puppeteer: {
        headless: true
    }
});

client.on('qr', async (qr) => {
    whatsappState = 'qr';

    try {
        // Convert WhatsApp's QR string to a Data URL.
        whatsappQr = await QRCode.toDataURL(qr);

        // console.log('New WhatsApp QR code generated.');
    } catch (error) {
        console.error('Failed to generate QR image:', error);
        whatsappQr = null;
    }
});

client.on('authenticated', () => {
    whatsappState = 'authenticated';
    console.log('WhatsApp authenticated.');
});

client.on('ready', () => {
    whatsappState = 'ready';
    console.log('WhatsApp client is ready!');
});

client.on('auth_failure', (message) => {
    whatsappState = 'auth_failure';
    console.error('WhatsApp authentication failure:', message);
});

client.on('disconnected', (reason) => {
    whatsappState = 'disconnected';
    console.log('WhatsApp disconnected:', reason);
});

client.on('message_ack', (msg, ack) => {
    if (ack >= MessageAck.ACK_SERVER) {
        ack_messages.push(msg);
    }
});

client.initialize();

/*
 * ---------------------------------------------------------
 * Helpers
 * ---------------------------------------------------------
 */

function normalizePhone(phone, prefix = '') {
    if (phone === null || phone === undefined) {
        return '';
    }

    let value = String(phone).trim();

    // Remove spaces, parentheses, hyphens, etc.
    value = value.replace(/[^\d+]/g, '');

    // If the number already has +, treat it as an international number.
    if (value.startsWith('+')) {
        return value.substring(1).replace(/\D/g, '');
    }

    // Remove leading zeroes when applying an international prefix.
    value = value.replace(/^0+/, '');

    const cleanPrefix = String(prefix)
        .replace(/\D/g, '');

    return cleanPrefix + value;
}

function renderMessage(template, row) {
    if (!template) {
        return '';
    }

    return template.replace(
        /\{([^{}]+)\}/g,
        (match, columnName) => {
            const value = row[columnName];

            return value === undefined || value === null
                ? ''
                : String(value);
        }
    );
}

function phoneToChatId(phone) {
    const normalized = normalizePhone(phone);

    if (!normalized) {
        throw new Error('Empty phone number');
    }

    return `${normalized}@c.us`;
}

function detectDelimiter(text) {
    const firstLines = text
        .replace(/^\uFEFF/, '')
        .split(/\r?\n/)
        .filter(line => line.trim())
        .slice(0, 10)
        .join('\n');

    const candidates = [
        { delimiter: ',', score: 0 },
        { delimiter: ';', score: 0 },
        { delimiter: '\t', score: 0 },
        { delimiter: '|', score: 0 }
    ];

    for (const candidate of candidates) {
        candidate.score = firstLines
            .split('\n')
            .reduce((total, line) => {
                return total + countDelimiterOutsideQuotes(
                    line,
                    candidate.delimiter
                );
            }, 0);
    }

    candidates.sort((a, b) => b.score - a.score);

    return candidates[0].score > 0
        ? candidates[0].delimiter
        : ',';
}

function countDelimiterOutsideQuotes(line, delimiter) {
    let count = 0;
    let insideQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];

        if (char === '"') {
            if (insideQuotes && line[i + 1] === '"') {
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (!insideQuotes && char === delimiter) {
            count++;
        }
    }

    return count;
}

/*
 * Small CSV parser supporting quoted fields and escaped quotes.
 */
function parseDelimitedText(text, delimiter) {
    text = text.replace(/^\uFEFF/, '');

    const rows = [];
    let row = [];
    let field = '';
    let insideQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];

        if (char === '"') {
            if (insideQuotes && text[i + 1] === '"') {
                field += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (char === delimiter && !insideQuotes) {
            row.push(field);
            field = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            if (char === '\r' && text[i + 1] === '\n') {
                i++;
            }

            row.push(field);
            field = '';

            if (row.some(value => value.trim() !== '')) {
                rows.push(row);
            }

            row = [];
        } else {
            field += char;
        }
    }

    // Last field/row
    row.push(field);

    if (row.some(value => value.trim() !== '')) {
        rows.push(row);
    }

    if (rows.length === 0) {
        return {
            headers: [],
            rows: []
        };
    }

    const headers = rows[0].map((header, index) => {
        const clean = String(header).trim();

        return clean || `Column ${index + 1}`;
    });

    const data = rows.slice(1).map(values => {
        const object = {};

        headers.forEach((header, index) => {
            object[header] = values[index] ?? '';
        });

        return object;
    });

    return {
        headers,
        rows: data
    };
}

function parseUploadedFile(buffer, originalName) {
    const extension = path.extname(originalName).toLowerCase();

    /*
     * Excel
     */
    if (extension === '.xlsx' || extension === '.xls') {
        const workbook = XLSX.read(buffer, {
            type: 'buffer',
            cellDates: true
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        if (!firstSheet) {
            throw new Error('The Excel file contains no worksheets.');
        }

        const rows = XLSX.utils.sheet_to_json(firstSheet, {
            defval: '',
            raw: false
        });

        const headers = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            defval: ''
        })[0] || [];

        return {
            headers: headers.map(String),
            rows
        };
    }

    /*
     * CSV/TXT
     */
    if (extension === '.csv' || extension === '.txt') {
        const text = buffer.toString('utf8');
        const delimiter = detectDelimiter(text);

        return parseDelimitedText(text, delimiter);
    }

    throw new Error(
        'Unsupported file type. Please use XLSX, XLS, CSV or TXT.'
    );
}

/*
 * Wait for WhatsApp's ACK_SERVER for a particular outgoing message.
 *
 * ACK_SERVER means whatsapp-web.js received the server acknowledgement.
 * It does NOT mean the recipient has read the message.
 */
// function waitForServerAck(message, timeoutMs = 15000) {
//     return new Promise((resolve) => {
//         let finished = false;

//         const finish = (result) => {
//             if (finished) {
//                 return;
//             }

//             finished = true;

//             clearTimeout(timeout);
//             client.off('message_ack', onAck);

//             resolve(result);
//         };

//         const onAck = (ackMessage, ack) => {
//             if (
//                 ackMessage.id &&
//                 message.id &&
//                 ackMessage.id._serialized === message.id._serialized &&
//                 ack >= MessageAck.ACK_SERVER
//             ) {
//                 finish(true);
//             }

//             if (ack === MessageAck.ACK_ERROR) {
//                 finish(false);
//             }
//         };

//         const timeout = setTimeout(() => {
//             finish(false);
//         }, timeoutMs);

//         client.on('message_ack', onAck);
//     });
// }

/*
 * Small delay between messages.
 *
 * This is intentionally conservative. Do not remove this and turn the
 * application into a high-speed bulk sender.
 */
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/*
 * ---------------------------------------------------------
 * API
 * ---------------------------------------------------------
 */

/*
 * Return WhatsApp QR code.
 */
app.get('/api/qr', (req, res) => {
    res.json({
        state: whatsappState,
        qr: whatsappQr
    });
});

/*
 * Return WhatsApp connection status.
 */
app.get('/api/status', (req, res) => {
    res.json({
        state: whatsappState
    });
});

/*
 * Get last opened file path.
 */
// app.get('/api/path', (req, res) => {
//     fs.readFile(path.join(__dirname, '/data/lastOpenedFile.txt'), { encoding: 'utf8' }, (err, data) => {
//         if (!err) {
//             res.json({
//                 path: data
//             });
//         } else {
//             console.log(`Error getting last opened file path: ${err}`)
//             res.json({
//                 path: ""
//             });
//         }
//     });
// });

/*
 * Parse an uploaded file.
 */
const multer = require('multer');

const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 25 * 1024 * 1024
    }
});

// app.post('/api/savePath', (req, res) => {
//     try {
//         fs.writeFile(path.join(__dirname, '/data/lastOpenedFile.txt'), `${req.body.path}`, { flag: 'w+' }, (err) => {
//         if (err) throw err;
//         });
//     } catch (error) {
//         console.log(`Error saving opened file path: ${error}`)
//     }
// });

app.post('/api/parse', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({
                error: 'No file uploaded.'
            });
        }

        const parsed = parseUploadedFile(
            req.file.buffer,
            req.file.originalname
        );

        if (parsed.headers.length === 0) {
            return res.status(400).json({
                error: 'The file does not contain a header row.'
            });
        }

        res.json({
            filename: req.file.originalname,
            headers: parsed.headers,
            rows: parsed.rows
        });
    } catch (error) {
        console.error('File parsing error:', error);

        res.status(400).json({
            error: error.message
        });
    }
});

/*
 * Send selected rows.
 *
 * The browser sends:
 *
 * {
 *   rows: [...],
 *   phoneColumn: "...",
 *   selectedIndexes: [...]
 * }
 */
let sending = false;

app.post('/api/send', async (req, res) => {
    if (sending) {
        return res.status(409).json({
            error: 'A sending operation is already in progress.'
        });
    }

    if (whatsappState !== 'ready') {
        return res.status(503).json({
            error: `WhatsApp is not ready. Current state: ${whatsappState}`
        });
    }

    const {
        rows,
        phoneColumn,
        messageTemplate,
        phonePrefix,
        selectedIndexes
    } = req.body;


    if (!Array.isArray(rows)) {
        return res.status(400).json({
            error: 'Invalid rows.'
        });
    }

    if (!phoneColumn) {
        return res.status(400).json({
            error: 'A phone-number column is required.'
        });
    }

    if (!Array.isArray(selectedIndexes)) {
        return res.status(400).json({
            error: 'Invalid selected rows.'
        });
    }

    sending = true;

    const outputRows = rows.map(row => ({
        ...row,
        finalMessage: ""
    }));

    const results = [];

    try {
        for (const index of selectedIndexes) {
            if (
                !Number.isInteger(index) ||
                index < 0 ||
                index >= rows.length
            ) {
                continue;
            }

            const row = rows[index];

            const phone = normalizePhone(
                row[phoneColumn],
                phonePrefix
            );

            let message;

            if (messageTemplate) {
                message = renderMessage(
                    messageTemplate,
                    row
                ).trim();
                outputRows[index].finalMessage = message;
            } else {
                continue;
            }


            if (!phone) {
                results.push({
                    index,
                    status: 'error',
                    error: 'Missing phone number'
                });

                continue;
            }

            if (!message) {
                results.push({
                    index,
                    status: 'error',
                    error: 'Message is empty'
                });

                continue;
            }

            try {
                const chatId = phoneToChatId(phone);

                /*
                 * Check whether WhatsApp knows this number.
                 *
                 * This avoids trying to send to an obviously invalid
                 * WhatsApp ID.
                 */
                const isRegistered =
                    await client.isRegisteredUser(chatId);

                if (!isRegistered) {
                    results.push({
                        index,
                        status: 'not_registered',
                        error: 'Phone number is not registered on WhatsApp'
                    });

                    continue;
                }

                const sentMessage = await client.sendMessage(
                    chatId,
                    message
                );

                /*
                 * Check for ACK_SERVER.
                 */
                let attempts = 0;
                let acknowledged = false;

                const interval = setInterval(() => {
                    attempts++;

                    acknowledged = ack_messages.includes(sentMessage);

                    if (acknowledged || attempts >= 10) {
                        clearInterval(interval);
                    }
                }, 1000);

                if (acknowledged) {
                    outputRows[index].wasSent = 1;

                    results.push({
                        index,
                        status: 'sent'
                    });
                } else {
                    outputRows[index].wasSent = 0;

                    results.push({
                        index,
                        status: 'failed',
                        error: 'No ACK_SERVER received'
                    });
                }

                /*
                 * Keep a delay between messages.
                 */
                await delay(1500);
            } catch (error) {
                console.error(
                    `Failed to send row ${index}:`,
                    error
                );

                results.push({
                    index,
                    status: 'error',
                    error: error.message
                });
            }
        }

        /*
         * Create XLSX output.
         */
        const worksheet = XLSX.utils.json_to_sheet(outputRows);
        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            'Results'
        );

        const outputBuffer = XLSX.write(workbook, {
            type: 'buffer',
            bookType: 'xlsx'
        });

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        res.setHeader(
            'Content-Disposition',
            'attachment; filename="whatsapp_results.xlsx"'
        );

        res.send(outputBuffer);
    } catch (error) {
        console.error('Sending error:', error);

        res.status(500).json({
            error: error.message,
            results
        });
    } finally {
        sending = false;
    }
});

/*
 * ---------------------------------------------------------
 * Start HTTP server
 * ---------------------------------------------------------
 */

app.listen(PORT, () => {
    console.log(
        `\nServer running at http://localhost:${PORT}\n`
    );
});