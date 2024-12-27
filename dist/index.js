#!/usr/bin/env node
"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    Object.defineProperty(o, k2, { enumerable: true, get: function() { return m[k]; } });
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const inquirer_1 = __importDefault(require("inquirer"));
const turndown_1 = __importDefault(require("turndown"));
const dotenv = __importStar(require("dotenv"));
const moment_1 = __importDefault(require("moment"));
const telegraf_1 = require("telegraf");
const proxy_agent_1 = __importDefault(require("proxy-agent"));
require("isomorphic-fetch");
const auth_provider_1 = require("./auth-provider");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const conf_1 = __importDefault(require("conf"));
const node_notifier_1 = __importDefault(require("node-notifier"));
dotenv.config();
moment_1.default.locale('ru');
const config = new conf_1.default();
const app = {
    id: process.env.APP_ID || '',
    secret: process.env.APP_SECRET || '',
    tenant: process.env.TENANT || 'common',
    scope: 'offline_access user.read mail.read'
};
(() => __awaiter(void 0, void 0, void 0, function* () {
    const authProvider = new auth_provider_1.AuthProvider(app);
    const client = microsoft_graph_client_1.Client.initWithMiddleware({ authProvider });
    const proxy = process.env.PROXY_URL;
    const telegram = new telegraf_1.Telegram(process.env.BOT_TOKEN || '', {
        agent: proxy ? new proxy_agent_1.default(proxy) : undefined
    });
    const me = yield client.api('/me').get();
    console.log(`Authorized as ${me.displayName} (${me.mail})`);
    let folderId = config.get('folderId');
    let chatId = config.get('chatId');
    let filterEmail = config.get('filterEmail');
    if (!folderId || !chatId || !filterEmail) {
        const mailFolders = yield client.api('/me/mailFolders').get();
        const answers = yield inquirer_1.default.prompt([
            {
                type: 'list',
                name: 'folderId',
                message: 'Please select the folder to forward:',
                choices: mailFolders.value.map(folder => ({
                    name: `${folder.displayName} (${folder.unreadItemCount}/${folder.totalItemCount})`,
                    value: folder.id
                })),
                default: config.get('folderId')
            },
            {
                type: 'input',
                name: 'chatId',
                message: 'Please input @channelname or chat id:',
                default: config.get('chatId')
            },
            {
                type: 'input',
                name: 'filterEmail',
                message: 'Please input email to get messages for:',
                default: config.get('filterEmail')
            }
        ]);
        folderId = answers.folderId;
        chatId = answers.chatId;
        filterEmail = answers.filterEmail;
        if (chatId.match(/^@/)) {
            chatId = (yield telegram.getChat(chatId)).id;
        }
        config.set('folderId', folderId);
        config.set('chatId', chatId);
        config.set('filterEmail', filterEmail);
    }
    const link = config.get('deltaLink') || config.get('nextLink') || `/me/mailFolders/${folderId}/messages/delta?$top=10&$orderby=receivedDateTime+desc`;
    const mail = yield client.api(link).get();
    console.debug(JSON.stringify(mail.value, undefined, 2));
    if (mail['@odata.deltaLink']) {
        config.set('deltaLink', mail['@odata.deltaLink']);
    }
    else {
        config.delete('deltaLink');
    }
    if (mail['@odata.nextLink']) {
        config.set('nextLink', mail['@odata.nextLink']);
    }
    else {
        config.delete('nextLink');
    }
    const turndownService = new turndown_1.default();
    const messages = mail.value.filter(m => {
        const { toRecipients, ccRecipients, bccRecipients } = m;
        const recepients = (toRecipients || [])
            .concat(ccRecipients || [])
            .concat(bccRecipients || [])
            .map(r => r.emailAddress || {})
            .map(a => a.address || '')
            .map(s => s.toLowerCase());
        return recepients.includes(filterEmail.toLowerCase());
    }).map(m => {
        const { subject, body, hasAttachments } = m;
        const content = body ? body.content || '' : '';
        return {
            subject: subject,
            body: turndownService.turndown(content)
                .replace(/^\s+/, '')
                .replace(/\s+$/, '')
                .replace(/\s*\n\s*/g, '\n'),
            attachments: hasAttachments || false
        };
    });
    messages.forEach((message) => __awaiter(void 0, void 0, void 0, function* () {
        yield telegram.sendMessage(chatId, `*${message.subject}*\n\n${message.body}${message.attachments ? '\n\n(Punggawa Bot 24/7 😕)' : ''}`, { parse_mode: 'Markdown' });
    }));
    yield telegram.setChatDescription(chatId, `Punggawa Bot 24/7: ${moment_1.default().format('lll')}`);
}))().catch(node_notifier_1.default.notify);
