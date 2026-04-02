const { 
    default: makeWASocket, 
    useMultiFileAuthState, 
    DisconnectReason,
    downloadMediaMessage,
    Browsers,
    fetchLatestBaileysVersion
} = require('@whiskeysockets/baileys');
const pino = require('pino');
const xlsx = require('xlsx');
const qrcode = require('qrcode-terminal');

async function connectToWhatsApp() {
    const { version, isLatest } = await fetchLatestBaileysVersion();
    console.log(`Using WA v${version.join('.')}, isLatest: ${isLatest}`);

    const { state, saveCreds } = await useMultiFileAuthState('auth_info_baileys');
    
    const socket = makeWASocket({
        version,
        auth: state,
        logger: pino({ level: 'silent' }),
        browser: Browsers.macOS('Desktop')
    });

    socket.ev.on('creds.update', saveCreds);

    socket.ev.on('connection.update', (update) => {
        const { connection, lastDisconnect, qr } = update;
        
        if (qr) {
            console.log('Scan the QR code below to connect to WhatsApp:');
            qrcode.generate(qr, { small: true });
        }

        if(connection === 'close') {
            const shouldReconnect = (lastDisconnect.error)?.output?.statusCode !== DisconnectReason.loggedOut;
            console.log('Connection closed due to ', lastDisconnect.error, ', reconnecting ', shouldReconnect);
            if(shouldReconnect) {
                connectToWhatsApp();
            }
        } else if(connection === 'open') {
            console.log('Opened connection, bot is ready to process Excel files!');
        }
    });

    socket.ev.on('messages.upsert', async (m) => {
        try {
            const message = m.messages[0];
            if (!message || !message.message) return;

            // Unpack message if wrapped by WhatsApp
            let msg = message.message;
            if (msg.ephemeralMessage) msg = msg.ephemeralMessage.message;
            if (msg.viewOnceMessage) msg = msg.viewOnceMessage.message;
            if (msg.viewOnceMessageV2) msg = msg.viewOnceMessageV2.message;
            if (msg.documentWithCaptionMessage) msg = msg.documentWithCaptionMessage.message;

            const docMessage = msg.documentMessage;
            if (docMessage) {
                const fileName = docMessage.fileName || '';
                const mimeType = docMessage.mimetype || '';
                
                if (!fileName.endsWith('.xlsx')) return;
                
                console.log(`Received Excel file: ${fileName} from ${message.key.remoteJid}`);
                
                // Reply to the user
                await socket.sendMessage(message.key.remoteJid, { text: 'Excel file received, processing...' });

                // Download the document
                const buffer = await downloadMediaMessage(
                    message,
                    'buffer',
                    { },
                    { 
                        logger: pino({ level: 'silent' }),
                        reuploadRequest: socket.updateMediaMessage
                    }
                );

                // Parse the Excel file
                const workbook = xlsx.read(buffer, { type: 'buffer' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                
                if (rawData.length === 0) {
                    await socket.sendMessage(message.key.remoteJid, { text: 'The Excel file is empty.' });
                    return;
                }

                console.log(`Found ${rawData.length} rows in the excel file.`);
                
                // Search for the header row dynamically
                let ledgerIdx = -1;
                let amountIdx = -1;
                let phoneIdx = -1;
                let dataStartIndex = -1;
                
                for (let i = 0; i < rawData.length; i++) {
                    const rowArray = rawData[i];
                    if (!Array.isArray(rowArray)) continue;
                    
                    for (let j = 0; j < rowArray.length; j++) {
                        const cellValue = String(rowArray[j] || '').trim().toLowerCase();
                        if (cellValue.includes('ledger name') || cellValue.includes('particulars')) {
                            ledgerIdx = j;
                        } else if (cellValue.includes('outstanding amount') || cellValue.includes('amount')) {
                            amountIdx = j;
                        } else if (cellValue.includes('phone') || cellValue.includes('contact') || cellValue.includes('mobile')) {
                            phoneIdx = j;
                        }
                    }
                    
                    // If we found the minimum required headers, this is the header row
                    if (ledgerIdx !== -1 && amountIdx !== -1) {
                        dataStartIndex = i + 1;
                        break;
                    }
                }

                if (dataStartIndex === -1) {
                    console.log('Unable to detect correct columns. First row elements:', rawData[0]);
                    await socket.sendMessage(message.key.remoteJid, { text: `Error: I couldn't find the 'Ledger Name' and 'Outstanding Amount' headers. Please ensure they exist. I see the first row as: ${JSON.stringify(rawData[0])}` });
                    return;
                }

                let previews = [];
                let sentCount = 0;
                let failedCount = 0;
                
                const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

                for (let i = dataStartIndex; i < rawData.length; i++) {
                    const rowArray = rawData[i];
                    if (!Array.isArray(rowArray)) continue;

                    const companyName = String(rowArray[ledgerIdx] || '').trim();
                    const outstandingAmount = String(rowArray[amountIdx] || '').trim();
                    let phoneNumber = phoneIdx !== -1 ? String(rowArray[phoneIdx] || '').trim() : '';

                    // Determine if it's the Total row or an empty row
                    if (!companyName || companyName.toLowerCase().includes('total') || !outstandingAmount) {
                        continue; // Skip empty rows or Total rows
                    }

                    const defaultMessage = `Dear Sir/Mam,\n\n${companyName}\nOutstanding Amount: ${outstandingAmount}\n\nThanks`;

                    if (phoneNumber) {
                        let originalPhone = phoneNumber;
                        phoneNumber = phoneNumber.replace(/\D/g, "");

                        // Extra fallback: if they already provided a number starting with 91, strip it so the 10-digit check passes
                        if (phoneNumber.length === 12 && phoneNumber.startsWith("91")) {
                            phoneNumber = phoneNumber.substring(2);
                        } else if (phoneNumber.length === 11 && phoneNumber.startsWith("0")) {
                            phoneNumber = phoneNumber.substring(1);
                        }

                        if (phoneNumber.length !== 10) {
                            console.log("❌ Skipping invalid number:", originalPhone, "->", phoneNumber);
                            continue;
                        }
                        
                        let formattedNumber = `91${phoneNumber}@s.whatsapp.net`;
                        
                        // Random delay between 15 and 30 seconds for all messages except the first
                        if (sentCount > 0 || failedCount > 0) {
                            const waitTimeSeconds = Math.floor(Math.random() * (30 - 15 + 1)) + 15;
                            console.log(`Waiting ${waitTimeSeconds} seconds before sending next message...`);
                            await sleep(waitTimeSeconds * 1000);
                        }

                        // Send message to the extracted phone number
                        try {
                            await socket.sendMessage(formattedNumber, { text: defaultMessage });
                            console.log(`✅ Sent message to ${formattedNumber} for ${companyName}`);
                            sentCount++;
                        } catch (err) {
                            console.log("❌ Failed for:", companyName, err.message);
                            failedCount++;
                        }
                    } else {
                        // Preview the message to the sender instead of actual sending
                        previews.push(defaultMessage);
                    }
                }

                if (previews.length > 0) {
                     let previewMsg = `No 'Phone Number' column found or numbers were blank. Here are the generated messages as a preview:\n\n`;
                     const maxPreviews = Math.min(previews.length, 5); // Limit to 5 previews to avoid spam
                     previewMsg += previews.slice(0, maxPreviews).join("\n--------------------\n");
                     
                     if (previews.length > maxPreviews) {
                         previewMsg += `\n\n...and ${previews.length - maxPreviews} more messages.`;
                     }
                     
                     previewMsg += `\n\nTo send to actual people, please include a 'Phone Number' column with country code.`;
                     
                     await socket.sendMessage(message.key.remoteJid, { text: previewMsg });
                } else if (sentCount > 0 || failedCount > 0) {
                    await socket.sendMessage(message.key.remoteJid, { text: `✅ Successfully sent ${sentCount} messages\n❌ Failed: ${failedCount}` });
                }
            }
        } catch (error) {
            console.error('Error handling message:', error);
        }
    });
}

connectToWhatsApp();
