import makeWASocket, {
    useMultiFileAuthState,
    DisconnectReason,
    downloadMediaMessage,
    Browsers,
    fetchLatestBaileysVersion
} from '@whiskeysockets/baileys';

import pino from 'pino';
import * as xlsx from 'xlsx';
import qrcode from 'qrcode-terminal';

async function connectToWhatsApp() {
    try {
        const { version } = await fetchLatestBaileysVersion();

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
                console.log('📱 Scan QR below:');
                qrcode.generate(qr, { small: true });
            }

            if (connection === 'close') {
                const shouldReconnect =
                    lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;

                console.log('❌ Disconnected. Reconnecting:', shouldReconnect);

                if (shouldReconnect) connectToWhatsApp();
            } else if (connection === 'open') {
                console.log('✅ Bot Connected Successfully!');
            }
        });

        socket.ev.on('messages.upsert', async (m) => {
            try {
                const message = m.messages[0];
                if (!message?.message) return;

                let msg = message.message;

                if (msg.ephemeralMessage) msg = msg.ephemeralMessage.message;
                if (msg.viewOnceMessage) msg = msg.viewOnceMessage.message;
                if (msg.viewOnceMessageV2) msg = msg.viewOnceMessageV2.message;
                if (msg.documentWithCaptionMessage) msg = msg.documentWithCaptionMessage.message;

                const doc = msg.documentMessage;
                if (!doc) return;

                const fileName = doc.fileName || '';
                if (!fileName.endsWith('.xlsx')) return;

                console.log(`📄 Received: ${fileName}`);

                await socket.sendMessage(message.key.remoteJid, {
                    text: '📄 Processing Excel file...'
                });

                const buffer = await downloadMediaMessage(
                    message,
                    'buffer',
                    {},
                    {
                        logger: pino({ level: 'silent' }),
                        reuploadRequest: socket.updateMediaMessage
                    }
                );

                const workbook = xlsx.read(buffer, { type: 'buffer' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

                if (!data.length) {
                    await socket.sendMessage(message.key.remoteJid, {
                        text: '❌ Excel is empty'
                    });
                    return;
                }

                let nameIdx = -1;
                let amountIdx = -1;
                let phoneIdx = -1;
                let startIndex = -1;

                for (let i = 0; i < data.length; i++) {
                    const row = data[i];
                    if (!Array.isArray(row)) continue;

                    for (let j = 0; j < row.length; j++) {
                        const cell = String(row[j] || '').toLowerCase();

                        if (cell.includes('ledger') || cell.includes('name')) nameIdx = j;
                        else if (cell.includes('amount')) amountIdx = j;
                        else if (cell.includes('phone') || cell.includes('mobile')) phoneIdx = j;
                    }

                    if (nameIdx !== -1 && amountIdx !== -1) {
                        startIndex = i + 1;
                        break;
                    }
                }

                if (startIndex === -1) {
                    await socket.sendMessage(message.key.remoteJid, {
                        text: '❌ Required columns not found'
                    });
                    return;
                }

                let sent = 0;
                let failed = 0;

                const sleep = (ms) => new Promise(r => setTimeout(r, ms));

                for (let i = startIndex; i < data.length; i++) {
                    const row = data[i];
                    if (!Array.isArray(row)) continue;

                    const name = String(row[nameIdx] || '').trim();
                    const amount = String(row[amountIdx] || '').trim();
                    let phone = phoneIdx !== -1 ? String(row[phoneIdx] || '') : '';

                    if (!name || !amount || name.toLowerCase().includes('total')) continue;

                    phone = phone.replace(/\D/g, '');

                    if (phone.length === 12 && phone.startsWith('91')) phone = phone.slice(2);
                    if (phone.length === 11 && phone.startsWith('0')) phone = phone.slice(1);

                    if (phone.length !== 10) {
                        console.log(`❌ Invalid: ${phone}`);
                        continue;
                    }

                    const jid = `91${phone}@s.whatsapp.net`;

                    const dateObj = new Date();
                    const currentDate = `${String(dateObj.getDate()).padStart(2, '0')}/${String(dateObj.getMonth() + 1).padStart(2, '0')}/${dateObj.getFullYear()}`;

                    const text = `Dear ${name},

📅 Date: ${currentDate}
➡️ Total Outstanding: ₹${amount}

Thank you 🙏`;

                    if (sent > 0) await sleep(15000);

                    try {
                        await socket.sendMessage(jid, { text });
                        console.log(`✅ Sent: ${jid}`);
                        sent++;
                    } catch (err) {
                        console.log(`❌ Failed: ${jid}`);
                        failed++;
                    }
                }

                await socket.sendMessage(message.key.remoteJid, {
                    text: `✅ Done\nSent: ${sent}\nFailed: ${failed}`
                });

            } catch (err) {
                console.error('❌ Error processing message:', err);
            }
        });

    } catch (err) {
        console.error('❌ Startup Error:', err);
    }
}

connectToWhatsApp();