// const ImapClient = require("emailjs-imap-client");
const ImapClient = require("imap");
import {ParsedMail} from "mailparser";
import {simpleParser} from "mailparser";
import {IServerInfo} from "./ServerInfo";

const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');



export interface ICallOptions {
    mailbox: string,
    id?: number;
}

export interface IMessage {
    id: string,
    date: string,
    from: string,
    subject: string,
    body?: string
}

export interface IMailbox {
    name: string,
    path: string
}

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

export class Worker {
    private static serverInfo: IServerInfo;
    // If modifying these scopes, delete token.json.
    private const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    private const TOKEN_PATH = 'token.json';

    constructor(inServerInfo: IServerInfo) {
        Worker.serverInfo = inServerInfo;
    }

    // Load client secrets from a local file.
    private readCredentials() {
        fs.readFile('credentials.json', (err, content) => {
            if (err) return console.log('Error loading client secret file:', err);
            // Authorize a client with credentials, then call the Gmail API.
            this.authorize(JSON.parse(content), this.listLabels);
        });
    }

    /**
     * Create an OAuth2 client with the given credentials, and then execute the
     * given callback function.
     * @param {Object} credentials The authorization client credentials.
     * @param {function} callback The callback to call with the authorized client.
     */
    private authorize(credentials, callback) {
        const {client_secret, client_id, redirect_uris} = credentials.installed;
        const oAuth2Client = new google.auth.OAuth2(
            client_id, client_secret, redirect_uris[0]);
    
        // Check if we have previously stored a token.
        fs.readFile(this.TOKEN_PATH, (err, token) => {
        if (err) return getNewToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        callback(oAuth2Client);
        });
    }

    /**
     * Get and store new token after prompting for user authorization, and then
     * execute the given callback with the authorized OAuth2 client.
     * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
     * @param {getEventsCallback} callback The callback for the authorized client.
     */
    private getNewToken(oAuth2Client, callback) {
        const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: this.SCOPES,
        });
        console.log('Authorize this app by visiting this url:', authUrl);
        const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
        });
        rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error retrieving access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(this.TOKEN_PATH, JSON.stringify(token), (err) => {
            if (err) return console.error(err);
            console.log('Token stored to', this.TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
        });
    }

    /**
     * Lists the labels in the user's account.
     *
     * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
     */
    private listLabels(auth) {
        const gmail = google.gmail({version: 'v1', auth});
        gmail.users.labels.list({
        userId: 'me',
        }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const labels = res.data.labels;
        if (labels.length) {
            console.log('Labels:');
            labels.forEach((label) => {
            console.log(`- ${label.name}`);
            });
        } else {
            console.log('No labels found.');
        }
        });
    }

    private async connectToServer(): Promise<any> {
        const client: any = new ImapClient.default(Worker.serverInfo.imap.host, Worker.serverInfo.imap.port, {
            auth: Worker.serverInfo.imap.auth
        });
    
        client.logLevel = client.LOG_LEVEL_NONE;

        client.onerror = (inError: Error) => {
            console.log(
                "IMAP.Worker.listMailboxes(): Connection error", inError);
        };
        
        await client.connect();
        return client;
    }

    public async listMailboxes(): Promise<IMailbox[]> {




        const client: any = await this.connectToServer();
        const mailboxes: any = await client.listMailboxes();
        await client.close();
        const finalMailboxes: IMailbox[] = [];
        const iterateChildren: Function = (inArray: any[]): void => {
            inArray.forEach((inValue: any) => {
                finalMailboxes.push({
                    name: inValue.name, path: inValue.path
                });
                iterateChildren(inValue.children);
            });
        };

        iterateChildren(mailboxes.children);

        return finalMailboxes;
    }

    public async listMessages(inCallOptions: ICallOptions): Promise<IMessage[]> {
        const client: any = await this.connectToServer();
        const mailbox: any = await client.selectMailbox(inCallOptions.mailbox);

        if (mailbox.exists === 0) {
            await client.close();
            return [];
        }

        const messages: any[] = await client.listMessages(
            inCallOptions.mailbox, "1:*", ["uid", "envelope"]
        );

        await client.close();

        const finalMessages: IMessage[] = [];

        messages.forEach((inValue: any) => {
            finalMessages.push({
                id: inValue.uid,
                date: inValue.envelope.date,
                from: inValue.envelope.from[0].address,
                subject: inValue.envelope.subject
            });
        });

        return finalMessages;
    }

    public async getMessageBody(inCallOptions: ICallOptions): Promise<string | undefined>{
        const client: any = await this.connectToServer();
        
        const messages: any[] = await client.listMessages(
            inCallOptions.mailbox, inCallOptions.id,
            [ "body[]" ], {byUid: true}
        );

        const parsed: ParsedMail = await simpleParser(messages[0]["body[]"]);
        await client.close();
        
        return parsed.text;
    }

    public async deleteMessage(inCallOptions: ICallOptions): Promise<any> {
        const client: any = await this.connectToServer();

        await client.deleteMessage(
            inCallOptions.mailbox, inCallOptions.id, {byUid: true}
        );

        await client.close();
    }
}
