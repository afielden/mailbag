import Mail from "nodemailer/lib/mailer";
import * as nodemailer from "nodemailer";
import { SendMailOptions, SentMessageInfo } from "nodemailer";
// const nodemailer = require("nodemailer");
import { serverInfo, IServerInfo } from "./ServerInfo";

export class Worker {
    private static serverInfo: IServerInfo;

    constructor(inServerInfo: IServerInfo) {
        Worker.serverInfo = inServerInfo;
    }

    public sendMessage(inOptions: SendMailOptions): Promise<string> {
        return new Promise((inResolve, inReject) => {
            const transport: Mail = nodemailer.createTransport(Worker.serverInfo.smtp);
            transport.sendMail(inOptions, (inError: Error | null, inInfo: SentMessageInfo) => {
                if (inError) {
                    inReject(inError);
                } else {
                    inResolve("ok");
                }
            })
        })
    }
}

