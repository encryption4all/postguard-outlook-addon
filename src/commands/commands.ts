/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import 'web-streams-polyfill'

import i18next from 'i18next'
import translationEN from '../../locales/en.json'
import translationNL from '../../locales/nl.json'

import { ComposeMail } from '@e4a/irmaseal-mail-utils/dist/composeMail'
import { createMimeMessage } from 'mimetext'
import {
  storeMailAsPlainLocally,
  IAttachmentContent,
  htmlBodyType,
  getItemRestId,
  isPostGuardEmail,
  newReadableStreamFromArray,
  getGlobal,
  getPostGuardHeaders,
  PKG_URL,
  showInfoMessage,
  Policy,
  checkLocalStorage,
  hashCon,
  getPublicKey
} from '../helpers/utils'

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

var globalEvent
var mailboxItem: Office.MessageCompose
var isEncryptMode: boolean = false
var isSignEmail: boolean = false

const EMAIL_ATTRIBUTE_TYPE = 'pbdf.sidn-pbdf.email.email'

const mod_promise = import('@e4a/pg-wasm') // require('@e4a/pg-wasm')
import { ISealOptions } from '@e4a/pg-wasm'
import jwtDecode, { JwtPayload } from 'jwt-decode'

const getLogger = require('webpack-log')
const encryptLog = getLogger({ name: 'PostGuard encrypt log' })
const decryptLog = getLogger({ name: 'PostGuard decrypt log' })

/**
 * The initialization function
 */
Office.initialize = () => {
  Office.onReady(() => {
    mailboxItem = Office.context.mailbox.item
    delete window.alert // assures alert works
    delete window.confirm // assures confirm works
    delete window.prompt // assures prompt works

    i18next.init({
      lng: Office.context.displayLanguage.toLowerCase().startsWith('nl')
        ? 'nl'
        : 'en',
      debug: true,
      resources: {
        en: {
          translation: translationEN
        },
        nl: {
          translation: translationNL
        }
      }
    })
  })
}

function encryptExtended(event: Office.AddinCommands.Event) {
  globalEvent = event
  showInfoMessage('Encrypting email with PostGuard extended')
  isEncryptMode = true
  showLoginPopup('/fallbackauthdialog.html')
}

function encryptExtendedSign(event: Office.AddinCommands.Event) {
  globalEvent = event
  showInfoMessage('Encrypting email with PostGuard extended + signatures')
  isSignEmail = true
  isEncryptMode = true
  showLoginPopup('/fallbackauthdialog.html')
}

/**
 * Gets the To of the current mail item
 * @returns The To as an array of string
 */
function getRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.to.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the Ccs of the current mail item
 * @returns The Ccs as an array of string
 */
function getCcRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.cc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the Bccs of the current mail item
 * @returns The Bccs as an array of string
 */
function getBccRecipientEmails(): Promise<string[]> {
  return new Promise(function (resolve) {
    mailboxItem.bcc.getAsync((recipients) => {
      let recipientMails = new Array()
      recipients.value.forEach((recipient) => {
        recipientMails.push(recipient.emailAddress.toLowerCase())
      })
      resolve(recipientMails)
    })
  })
}

/**
 * Gets the body of the current mail item
 * @returns The body
 */
async function getMailBody(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Body async failed')
      } else {
        let returnBody: string = asyncResult.value
        if (returnBody !== '') resolve(returnBody)
        else reject('Please add text to the body in the email')
      }
    })
  })
}

/**
 * Gets the subject from the current mail item
 * @returns The subject
 */
async function getMailSubject(): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Subject async failed')
      } else {
        const subject: string = asyncResult.value
        if (subject !== '') resolve(subject)
        else reject('Please add a subject to the email')
      }
    })
  })
}

/**
 * Collects all metadata and content of attachment belonging to the current mail item
 * @returns Array of IAttachmentContent objects
 */
async function getMailAttachments(): Promise<IAttachmentContent[]> {
  return new Promise(function (resolve, reject) {
    mailboxItem.getAttachmentsAsync(async (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Attachments async failed')
      } else {
        let attachmentsArray = []
        let content = ''
        for (var i = 0; i < asyncResult.value.length; i++) {
          var attachment = asyncResult.value[i]
          content = await getMailAttachmentContent(attachment.id)
          attachmentsArray.push({
            filename: attachment.name,
            content: content,
            isInline: attachment.isInline,
            id: attachment.id
          })
        }
        resolve(attachmentsArray)
      }
    })
  })
}

/**
 * Returns promise containing content of the attachment
 * @param attachmentId ID of the attachment
 * @returns Promise containing content of the attachment
 */
async function getMailAttachmentContent(attachmentId: string): Promise<string> {
  return new Promise(function (resolve, reject) {
    mailboxItem.getAttachmentContentAsync(attachmentId, (asyncResult) => {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        reject('Attachment content async failed')
      } else {
        if (asyncResult.value.content.length > 0) {
          resolve(asyncResult.value.content)
        } else
          reject('No attachment content in attachment with id' + attachmentId)
      }
    })
  })
}

async function getSigningKeys(jwt: string, keyRequest?: any): Promise<any> {
  const url = `${PKG_URL}/v2/irma/sign/key`
  return fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${jwt}`,
      'X-Postguard-Client-Version': getPostGuardHeaders(),
      'content-type': 'application/json'
    },
    body: JSON.stringify(keyRequest)
  })
    .then((r) => r.json())
    .then((json) => {
      if (json.status !== 'DONE' || json.proofStatus !== 'VALID')
        throw new Error('session not DONE and VALID')
      return { pubSignKey: json.pubSignKey, privSignKey: json.privSignKey }
    })
}

/**
 * Encrypts and sends the email
 * @param token The authentication token for the Graph API
 * @param policy The encryption access policy
 * @param signingJwt Signing JWT
 */
async function encryptAndSendEmail(token, policy: Policy = null, jwt: any) {
  const pk = await getPublicKey()

  const [mod] = await Promise.all([mod_promise])

  const sender = Office.context.mailbox.userProfile.emailAddress

  const timestamp = Math.round(Date.now() / 1000)

  const recipientEmails: string[] = await getRecipientEmails()
  const ccRecipientEmails: string[] = await getCcRecipientEmails()
  const bccRecipientEmails: string[] = await getBccRecipientEmails()

  // if BCC recipients available , abort operations
  if (bccRecipientEmails.length > 0) {
    bccMsgAndDialog()
    return
  }

  const allRecipientsCount = recipientEmails.length + ccRecipientEmails.length

  if (allRecipientsCount == 0) {
    throw 'Please add recipients to the email!'
  }

  let allPolicies = {}

  if (policy === null) {
    const policies = recipientEmails.reduce((total, recipient) => {
      total[recipient] = {
        ts: timestamp,
        con: [{ t: EMAIL_ATTRIBUTE_TYPE, v: recipient }]
      }
      return total
    }, {})

    const ccPolicies = ccRecipientEmails.reduce((total, recipient) => {
      total[recipient] = {
        ts: timestamp,
        con: [{ t: EMAIL_ATTRIBUTE_TYPE, v: recipient }]
      }
      return total
    }, {})

    allPolicies = { ...policies, ...ccPolicies }
  } else {
    for (var id in policy) {
      allPolicies[id] = {
        ts: timestamp,
        con: policy[id]
      }
    }
  }

  encryptLog.info('Encrypting using the following policies: ', allPolicies)

  let mailBody = await getMailBody()

  const mailSubject = await getMailSubject()
  encryptLog.info('Mail subject: ', mailSubject)

  let attachments: IAttachmentContent[] = await getMailAttachments()

  // Use createMimeMessage to create inner MIME mail
  const msg = createMimeMessage()
  msg.setSender(sender)
  msg.setSubject(mailSubject)

  recipientEmails.length > 0 && msg.setRecipient(recipientEmails)
  ccRecipientEmails.length > 0 && msg.setCc(ccRecipientEmails)
  bccRecipientEmails.length > 0 && msg.setBcc(bccRecipientEmails)

  // ComposeMail only used for outer mail
  const composeMail = new ComposeMail()
  composeMail.setSubject('PostGuard Encrypted Email')
  composeMail.setSender(sender)

  recipientEmails.forEach((recipientEmail) => {
    composeMail.addRecipient(recipientEmail)
  })
  ccRecipientEmails.forEach((recipientEmail) => {
    composeMail.addCcRecipient(recipientEmail)
  })
  bccRecipientEmails.forEach((recipientEmail) => {
    composeMail.addBccRecipient(recipientEmail)
  })

  let hasAttachments: boolean = false

  if (attachments !== undefined) {
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]

      if (!attachment.isInline) {
        hasAttachments = true
        const input = new TextEncoder().encode(attachment.content)
        encryptLog.info('Attachment bytes length: ', input.byteLength)
        msg.setAttachment(
          attachment.filename,
          'application/octet-stream',
          attachment.content
        )
      } else {
        // replace inline image in body
        const imageContentIDToReplace = `cid:${attachment.filename}@.*"`
        const regex = new RegExp(imageContentIDToReplace, 'g')
        mailBody = mailBody.replace(
          regex,
          `data:image;base64,${attachment.content}"`
        )
      }
    }
  }

  encryptLog.info('Mailbody: ', mailBody)
  msg.setMessage('text/html', mailBody)

  // encrypt inner MIME mail
  const innerMail = msg.asRaw()
  const plainBytes: Uint8Array = new TextEncoder().encode(innerMail)
  const readable = newReadableStreamFromArray(plainBytes)
  let ct = new Uint8Array(0)
  const writable = new WritableStream({
    write(chunk) {
      ct = new Uint8Array([...ct, ...chunk])
    }
  })
  const pubSignId = g.pubSignId
  const privSignId = g.privSignId

  const { pubSignKey, privSignKey } = await getSigningKeys(jwt, {
    pubSignId,
    privSignId
  })

  const sealOptions: ISealOptions = {
    policy: allPolicies,
    pubSignKey,
    ...(privSignKey && { privSignKey })
  }

  encryptLog.info('Sealing with options: ', sealOptions)

  await mod.sealStream(pk, sealOptions, readable, writable)

  // get outer mail to send email via Graph API
  composeMail.setPayload(ct)
  const outerMail = composeMail.getMimeMail()
  const message = Buffer.from(outerMail).toString('base64')
  const sendMessageUrl = 'https://graph.microsoft.com/v1.0/me/sendMail'
  encryptLog.info('Trying to send email via ', sendMessageUrl)

  $.ajax({
    type: 'POST',
    contentType: 'text/plain',
    url: sendMessageUrl,
    data: message,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function () {
      encryptLog.info(
        'Sendmail success, now trying to store mail locally, and clear mail'
      )

      let jsonInnerMail = {
        sender: { emailAddress: { address: sender } },
        subject: mailSubject,
        body: {
          contentType: htmlBodyType,
          content: mailBody
        },
        toRecipients: recipientEmails.map((recipient) => {
          return {
            emailAddress: { address: recipient }
          }
        }),
        ccRecipients: ccRecipientEmails.map((recipient) => {
          return {
            emailAddress: { address: recipient }
          }
        }),
        bccRecipients: bccRecipientEmails.map((recipient) => {
          return {
            emailAddress: { address: recipient }
          }
        }),
        hasAttachments: hasAttachments
      }

      storeMailAsPlainLocally(
        token,
        jsonInnerMail,
        attachments,
        'PostGuard Sent'
      )

      clearCurrentEmail(attachments)

      successMsgAndDialog()
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    encryptLog.error('Ajax error send mail: ', data)
    throw 'Error when sending mail, please try again or contact administrator'
  })
}

/**
 * Adds info message to email and shows dialog that PostGuard does not support BCC yet
 */
function bccMsgAndDialog() {
  showInfoMessage(
    'PostGuard does not support BCCs. Please remove BCCs, or send the mail unencrypted.'
  )
}

/**
 * Adds info message to email and shows dialog that message is successfully encrypted and send
 */
function successMsgAndDialog() {
  showInfoMessage('Successfully signed, encrypted and sent')
  var fullUrl =
    'https://' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/success.html'

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 30 },
    // eslint-disable-next-line no-unused-vars
    function (result) {
      encryptLog.info('Successdialog has initialized.')
    }
  )
}

/**
 * Clears the current mail, which means that to, cc, bcc, subject, body and attachments are removed
 * @param attachments The attachments of the mail
 */

function clearCurrentEmail(attachments) {
  mailboxItem.to.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.cc.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.bcc.setAsync([], function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.subject.setAsync('', function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing subject failed with error: ' + asyncResult.error.message
      )
    }
  })
  mailboxItem.body.setAsync('', (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      encryptLog.error(
        'Clearing body failed with error: ' + asyncResult.error.message
      )
    }
  })
  if (attachments !== undefined) {
    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i]
      mailboxItem.removeAttachmentAsync(attachment.id, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          encryptLog.error(
            'Clearing attachment failed with error: ' +
              asyncResult.error.message
          )
        }
      })
    }
  }
}

var signingDialog
var recipientAttributeDialog
var signingAttributeDialog

/**
 * Shows dialog to add attributes for each recipient
 * @param accessToken Token to access graph API
 */
async function chooseRecipientAcessPolicies(accessToken: string) {
  const recipientEmails: string[] = await getRecipientEmails()
  const ccRecipientEmails: string[] = await getCcRecipientEmails()
  const bccRecipientEmails: string[] = await getBccRecipientEmails()

  // if BCC recipients available , abort operations
  if (bccRecipientEmails.length > 0) {
    bccMsgAndDialog()
    return
  }

  const allRecipientsCount = recipientEmails.length + ccRecipientEmails.length

  if (allRecipientsCount == 0) {
    throw 'Please add recipients to the email!'
  }

  const recipientsStringified = JSON.stringify(
    recipientEmails.concat(ccRecipientEmails)
  )

  const b64Recipients = Buffer.from(recipientsStringified).toString('base64')
  const b64Token = Buffer.from(accessToken).toString('base64')

  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/attributes.html' +
    '?recipients=' +
    b64Recipients +
    '&token=' +
    b64Token

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 20 },
    function (result) {
      encryptLog.info('Recipients attributedialog has initialized.')
      recipientAttributeDialog = result.value
      recipientAttributeDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processRecipientAccessPolicyMessage
      )
    }
  )
}

/**
 *  This handler responds to the success or failure message that the pop-up dialog receives from the attribute form handler.
 * @param arg The arg object passed from the dialog
 */

async function processRecipientAccessPolicyMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)
  recipientAttributeDialog.close()

  if (messageFromDialog.status === 'success') {
    g.accessPolicy = messageFromDialog.result.policy
    const recipients = Object.keys(g.accessPolicy)

    Office.context.mailbox.item.to.setAsync(recipients, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // if global signingJwt is null, it means there is no JWT available for that sigining policy
        if (g.signingJwt === null) {
          signMessageDialog(messageFromDialog.result.accessToken)
        } else {
          // if JWT is available, directly send and decrpt
          encryptAndSendEmail(
            messageFromDialog.result.accessToken,
            g.accessPolicy,
            g.signingJwt
          ).catch((err) => {
            encryptLog.error(err)
            showInfoMessage(err)
            globalEvent.completed()
          })
        }
      } else {
        throw 'Could not set recipients in compose window'
      }
    })
  } else {
    encryptLog.error(
      'Error: ',
      JSON.stringify(messageFromDialog.error.toString())
    )
  }
}

async function signMessageDialog(accessToken: String) {
  const b64Token = Buffer.from(accessToken).toString('base64')
  const b64Policy = Buffer.from(JSON.stringify(g.signingpolicy)).toString(
    'base64'
  )
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/sign.html' +
    '?signingpolicy=' +
    b64Policy +
    '&token=' +
    b64Token

  encryptLog.info('signing message url: ', fullUrl)

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 20, promptBeforeOpen: false },
    function (result) {
      encryptLog.info('Signing attributedialog has initialized.')
      signingDialog = result.value
      signingDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processSignMsgMessage // receive used policy and store in object as it is later used again
      )
    }
  )
}

var loginDialog

/**
 *  This handler responds to the success or failure message that the pop-up dialog receives from the identity provider and access token provider.
 * @param arg The arg object passed from the dialog
 */
async function processAuthMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)
  const accessToken = messageFromDialog.result.accessToken

  encryptLog.info(`After auth: ${JSON.stringify(messageFromDialog)}`)

  if (messageFromDialog.status === 'success') {
    loginDialog.close()
    if (isEncryptMode) {
      g.msgFunc = showInfoMessage

      // show attribute form to choose sign policy
      if (isSignEmail) {
        chooseSigningPolicy(accessToken).catch((err) => {
          encryptLog.error(err)
          showInfoMessage(err)
          globalEvent.completed()
        })
      } else {
        chooseRecipientAcessPolicies(accessToken).catch((err) => {
          encryptLog.error(err)
          showInfoMessage(err)
          globalEvent.completed()
        })
      }
    } else {
      showDecryptPopup(accessToken)
    }
  } else {
    // Something went wrong with authentication or the authorization of the web application... try again
    encryptLog.error('Error: ', JSON.stringify(messageFromDialog))
    globalEvent.completed()
  }
}

async function processSigningPolicyMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)

  signingAttributeDialog.close()

  if (messageFromDialog.status === 'success') {
    const accessToken = messageFromDialog.result.accessToken

    encryptLog.info(
      `Signing policy: ${JSON.stringify(messageFromDialog.result.policy)}`
    )

    const pubSignId = [
      {
        t: EMAIL_ATTRIBUTE_TYPE,
        v: Office.context.mailbox.userProfile.emailAddress
      }
    ]

    const privSignId = messageFromDialog.result.policy[
      Office.context.mailbox.userProfile.emailAddress
    ].filter(({ t }) => t !== EMAIL_ATTRIBUTE_TYPE)

    encryptLog.info(`Private signing policy: ${JSON.stringify(privSignId)}`)

    g.pubSignId = pubSignId
    g.privSignId = privSignId
    g.signingpolicy = [...pubSignId, ...(privSignId ? privSignId : [])]
    g.signingJwt = await checkLocalStorage(g.signingpolicy).catch((e) => null)

    encryptLog.info(`Signing JWT in storage: ${g.signingJwt}`)

    if (messageFromDialog.status === 'success') {
      chooseRecipientAcessPolicies(accessToken).catch((err) => {
        encryptLog.error(err)
        showInfoMessage(err)
        globalEvent.completed()
      })
    }
  } else {
    const err = JSON.stringify(messageFromDialog.error.toString())
    encryptLog.error('Error: ', err)
    showInfoMessage(err)
    globalEvent.completed()
  }
}

async function processSignMsgMessage(arg) {
  let messageFromDialog = JSON.parse(arg.message)
  signingDialog.close()

  if (messageFromDialog.status === 'success') {
    const accessToken = messageFromDialog.result.accessToken
    encryptLog.info(`Signing keys: ${JSON.stringify(messageFromDialog)}`)
    const signingJwt = messageFromDialog.result.jwt

    // store in local storage
    const hashPolicy = await hashCon(g.signingpolicy)
    const decoded = jwtDecode<JwtPayload>(signingJwt)
    window.localStorage.setItem(
      `jwt_${hashPolicy}`,
      JSON.stringify({ jwt: signingJwt, exp: decoded.exp })
    )

    encryptAndSendEmail(accessToken, g.accessPolicy, signingJwt).catch(
      (err) => {
        encryptLog.error(err)
        showInfoMessage(err)
        globalEvent.completed()
      }
    )
  } else {
    const err = JSON.stringify(messageFromDialog.error.toString())
    encryptLog.error('Error: ', err)
    showInfoMessage(err)
    globalEvent.completed()
  }
}

/**
 * Opens dialog to choose a signing policy
 * @param accessToken The access token
 */
async function chooseSigningPolicy(accessToken: String) {
  const b64Sender = Buffer.from(
    Office.context.mailbox.userProfile.emailAddress
  ).toString('base64')
  const b64Token = Buffer.from(accessToken).toString('base64')

  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/attributes.html' +
    '?sender=' +
    b64Sender +
    '&token=' +
    b64Token

  encryptLog.info('signing policy url: ', fullUrl)

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 40, width: 20 },
    function (result) {
      encryptLog.info('Signing attributedialog has initialized.')
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          chooseSigningPolicy(accessToken)
        }
      } else {
        signingAttributeDialog = result.value
        signingAttributeDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processSigningPolicyMessage // receive used policy and store in object as it is later used again
        )
      }
    }
  )
}

/**
 * Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
 * @param url The HTML url pointing to the login popup
 */

function showLoginPopup(url: string) {
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    url +
    '?currentAccountMail=' +
    Office.context.mailbox.userProfile.emailAddress

  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 60, width: 30, promptBeforeOpen: false },
    function (result) {
      encryptLog.info('Logindialog has initialized.')
      loginDialog = result.value
      loginDialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        processAuthMessage
      )
    }
  )
}

// eslint-disable-next-line no-unused-vars
var decryptDialog: Office.Dialog

/**
 * Decrypts current message via popup
 * @param event The AddinCommands Event
 */
function decrypt(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Decrypting email with PostGuard',
    icon: 'Icon.80x80',
    persistent: true
  }

  globalEvent = event

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )

  if (isPostGuardEmail()) {
    showLoginPopup('/fallbackauthdialog.html')
  } else {
    showInfoMessage(i18next.t('displayMessageNoPostGuardMail'))
  }
}

/**
 * Displays the decrypt popup
 * @param token Authentication token for Graph API
 */
function showDecryptPopup(token: string) {
  const b64 = Buffer.from(token).toString('base64')

  const fullUrl =
    'https://' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/decrypt.html' +
    '?token=' +
    b64 +
    '&mailid=' +
    getItemRestId() +
    '&recipient=' +
    Office.context.mailbox.userProfile.emailAddress +
    '&attachmentid=' +
    Office.context.mailbox.item.attachments[0].id
      .replace('/', '-')
      .replace('+', '_') +
    '&sender=' +
    Office.context.mailbox.item.sender.emailAddress

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 80, width: 30 },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        if (result.error.code === 12007) {
          showDecryptPopup(token) // Recursive call
        } else {
          decryptLog.error('Decryptdialog error: ', result.error)
          showInfoMessage(result.error.message)
        }
      } else {
        decryptLog.info('Decryptdialog has initialized, ', result)
        decryptDialog = result.value
        decryptDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          processDecryptMessage
        )
      }
    }
  )
}

/**
 * Event handler function receiving message form decrypt dialog
 * Passing message on to showInfoMessage function
 * @param arg object passed from decrypt dialog.
 */
function processDecryptMessage(arg) {
  showInfoMessage(arg.message)
}

// the add-in command functions need to be available in global scope
const g = getGlobal() as any
g.encrypt = encryptExtended
g.encryptExt = encryptExtended
g.encryptExtSign = encryptExtendedSign
g.decrypt = decrypt
