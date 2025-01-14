#!/usr/bin/env node

import inquirer from 'inquirer'
import TurndownService from 'turndown'
import * as dotenv from 'dotenv'
import moment from 'moment'
import { Telegram } from 'telegraf'
import ProxyAgent from 'proxy-agent'
import 'isomorphic-fetch'
import { AppData, AuthProvider } from './auth-provider'
import { MailFolder, Message, User } from '@microsoft/microsoft-graph-types'
import { Client } from '@microsoft/microsoft-graph-client'
import { Agent } from 'https'
import Conf from 'conf'
import notifier from 'node-notifier'

dotenv.config()

moment.locale('ru')

const config = new Conf()

const app: AppData = {
  id: process.env.APP_ID || '',
  secret: process.env.APP_SECRET || '',
  tenant: process.env.TENANT || 'common',
  scope: 'offline_access user.read mail.read'
};

(async (): Promise<void> => {
  const authProvider: AuthProvider = new AuthProvider(app)
  const client: Client = Client.initWithMiddleware({ authProvider })

  const proxy = process.env.PROXY_URL
  const telegram = new Telegram(process.env.BOT_TOKEN || '', {
    agent: proxy ? new ProxyAgent(proxy) as unknown as Agent : undefined
  })

  const me: User = await client.api('/me').get()
  console.log(`Authorized as ${me.displayName} (${me.mail})`)

  let folderId = config.get('folderId')
  let chatId = config.get('chatId')
  let filterEmail = config.get('filterEmail')

  if (!folderId || !chatId || !filterEmail) {
    const mailFolders: { value: MailFolder[] } = await client.api('/me/mailFolders').get()

    const answers = await inquirer.prompt([
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
    ])

    folderId = answers.folderId
    chatId = answers.chatId
    filterEmail = answers.filterEmail

    if (chatId.match(/^@/)) {
      chatId = (await telegram.getChat(chatId)).id
    }

    config.set('folderId', folderId)
    config.set('chatId', chatId)
    config.set('filterEmail', filterEmail)
  }

  const link = config.get('deltaLink') || config.get('nextLink') || `/me/mailFolders/${folderId}/messages/delta?$top=10&$orderby=receivedDateTime+desc`

  const mail: {
    value: Message[];
    '@odata.nextLink'?: string;
    '@odata.deltaLink'?: string;
  } =
    await client.api(link).get()

  console.debug(JSON.stringify(mail.value, undefined, 2))

  if (mail['@odata.deltaLink']) {
    config.set('deltaLink', mail['@odata.deltaLink'])
  } else {
    config.delete('deltaLink')
  }

  if (mail['@odata.nextLink']) {
    config.set('nextLink', mail['@odata.nextLink'])
  } else {
    config.delete('nextLink')
  }

  const turndownService = new TurndownService()

  const messages = mail.value.filter(m => {
    const { toRecipients, ccRecipients, bccRecipients } = m
    const recepients: string[] = (toRecipients || [])
      .concat(ccRecipients || [])
      .concat(bccRecipients || [])
      .map(r => r.emailAddress || {})
      .map(a => a.address || '')
      .map(s => s.toLowerCase())
    return recepients.includes(filterEmail.toLowerCase())
  }).map(m => {
    const { subject, body, hasAttachments } = m
    const content: string = body ? body.content || '' : ''

    return {
      subject: subject,
      body: turndownService.turndown(content)
        .replace(/^\s+/, '')
        .replace(/\s+$/, '')
        .replace(/\s*\n\s*/g, '\n'),
      attachments: hasAttachments || false
    }
  })

  messages.forEach(async message => {
    await telegram.sendMessage(chatId, `*${message.subject}*\n\n${message.body}${message.attachments ? '\n\n(Punggawa Bot 24/7 😕)' : ''}`, { parse_mode: 'Markdown' })
  })
  await (telegram as any).setChatDescription(chatId, `Punggawa Bot 24/7: ${moment().format('lll')}`)
})().catch(notifier.notify)
