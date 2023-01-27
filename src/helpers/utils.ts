/* global $, Office */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'

export const htmlBodyType: MicrosoftGraph.BodyType = 'html'

import * as getLogger from 'webpack-log'
const utilLog = getLogger({ name: 'PostGuard util log' })

/**
 * Interface to store attachment metadata and content
 */
export interface IAttachmentContent {
  filename: string
  content: string
  isInline: boolean
  id: string
}

/**
 * Handles an ajax error
 * @param $xhr The error
 */
function handleAjaxError($xhr) {
  var data = $xhr.responseJSON
  utilLog.info('Ajax error: ', data)
  setEventError()
}

/**
 * Replaces the mail body
 * @param token The authentication token for the Graph aPI
 * @param item The mail item
 * @param body The body of the mail
 * @param attachments The decrypted attachments to be added to the mail
 */
export function replaceMailBody(
  token: string,
  item: any,
  body: string,
  attachments: IAttachmentContent[]
) {
  const itemId = getItemRestId()
  const messageUrl = `https://graph.microsoft.com/v1.0/me/messages/${itemId}`
  const payload = {
    body: {
      contentType: htmlBodyType,
      content: body
    }
  }
  $.ajax({
    type: 'PATCH',
    contentType: 'application/json',
    url: messageUrl,
    data: JSON.stringify(payload),
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('PATCH message success: ', success)
      removeAttachment(token, itemId, item.attachments[0].id, attachments)
    }
  }).fail(handleAjaxError)
}

/**
 * Removes an attachment
 * @param token The authentication token for the Graph aPI
 * @param itemId The id of the mail with the attachment
 * @param attachmentId The attachment id
 * @param attachments The decrypted attachments added to the mail later
 */
export function removeAttachment(
  token: string,
  itemId: string,
  attachmentId: string,
  attachments: IAttachmentContent[]
) {
  const attachmentUrl = `https://graph.microsoft.com/v1.0/me/messages/${itemId}/attachments/${attachmentId}`
  utilLog.info(`AttachmentURL: ${attachmentUrl}`)
  $.ajax({
    type: 'DELETE',
    url: attachmentUrl,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('DELETE attachment success: ', success)
      attachments.forEach((attachment) => {
        addAttachment(token, itemId, attachment)
      })
    }
  }).fail(handleAjaxError)
}

/**
 *
 * @param token The authentication token for the Graph API
 * @param messageId The id of the mail
 * @param attachment The attachment
 */
function addAttachment(
  token: string,
  messageId: string,
  attachment: IAttachmentContent
) {
  const createAttachmentUrl = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`

  const jsonAttachment = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: attachment.filename,
    contentBytes: attachment.content,
    isInline: attachment.isInline
  }

  $.ajax({
    type: 'POST',
    contentType: 'application/json',
    url: createAttachmentUrl,
    data: JSON.stringify(jsonAttachment),
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('Add attachment success: ', success)
    }
  }).fail(handleAjaxError)
}

/**
 * Stores decrypted mail locally by first checking if mail folder with folderName already exists
 * @param token The authentication token for the Graph API
 * @param innerMail The inner mail
 * @param attachments The attachments passed to another function
 * @param folderName The name of the folder
 */
export function storeMailAsPlainLocally(
  token: string,
  innerMail: MicrosoftGraph.Message,
  attachments: IAttachmentContent[],
  folderName: string
) {
  const mailFoldersUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders'
  $.ajax({
    type: 'GET',
    url: mailFoldersUrl,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('MailFolders: ', success)
      let folderFound = false
      success.value.forEach((folder) => {
        if (!folderFound && folder.displayName === folderName) {
          folderFound = true
          utilLog.info('Folder exists with id ', folder.id)
          storeInnerMail(folder.id, innerMail, token, attachments)
        }
      })
      if (!folderFound) {
        utilLog.info('Folder not found, creating ...')
        createPostGuardMailFolder(token, innerMail, attachments, folderName)
      }
    }
  }).fail(handleAjaxError)
}

/**
 * Creates the postguard mail folder
 * @param token The authentication token for the Graph API
 * @param innerMail The inner mail
 * @param attachments The attachments passed to another function
 * @param folderName The name of the folder
 */
function createPostGuardMailFolder(token, innerMail, attachments, folderName) {
  const createMailFoldersUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders'
  const payload = {
    displayName: folderName,
    isHidden: false
  }

  $.ajax({
    type: 'POST',
    contentType: 'application/json',
    url: createMailFoldersUrl,
    data: JSON.stringify(payload),
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('Created mailfolder succesfully!')
      storeInnerMail(success.id, innerMail, token, attachments)
    }
  }).fail(handleAjaxError)
}

/**
 * Store the inner mail
 * @param folderId The folder id the mail is stored in
 * @param innerMail The inner mail
 * @param token The authentication token for the Graph API
 * @param attachments The attachments passed to another function
 */
function storeInnerMail(
  folderId,
  innerMail,
  token,
  attachments: IAttachmentContent[]
) {
  const createMessageUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages`

  $.ajax({
    type: 'POST',
    contentType: 'application/json',
    url: createMessageUrl,
    data: JSON.stringify(innerMail),
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('Createmail success: ', success)
      attachments.forEach((attachment) => {
        storeAttachment(folderId, success.id, token, attachment)
      })
    }
  }).fail(handleAjaxError)
}

/**
 * Adds an attachment to a mail
 * @param folderId Id of the folder the mail is stored in
 * @param messageId Id of the mail
 * @param token Authentication token for Graph API
 * @param attachment The attachment to be stored
 */
function storeAttachment(
  folderId,
  messageId,
  token,
  attachment: IAttachmentContent
) {
  const createAttachmentUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages/${messageId}/attachments`

  const jsonAttachment = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: attachment.filename,
    contentBytes: attachment.content,
    isInline: attachment.isInline
  }

  $.ajax({
    type: 'POST',
    contentType: 'application/json',
    url: createAttachmentUrl,
    data: JSON.stringify(jsonAttachment),
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      utilLog.info('Create attachment success: ', success)
    }
  }).fail(handleAjaxError)
}

/**
 * Sets event error for encryption
 */
export function setEventError() {
  const msg = 'PostGuard error, please try again or contact your administrator.'
  // if mailbox is available, current context is the main window, and not a dialog
  if (Office.context.mailbox !== undefined) {
    const message: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: msg
    }
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      'action',
      message
    )
  } else {
    Office.context.ui.messageParent(msg)
  }
}

/**
 * Determines the item id of the current mail item
 * @returns The item id of the current mail item
 */
export function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    )
  }
}

/**
 * Calculates a hash for a message
 * @param message The message
 * @returns The hash
 */
export async function hashString(message: string): Promise<string> {
  const msgArray = new TextEncoder().encode(message)
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgArray)
  const hashArray = Array.from(new Uint8Array(hashBuffer))
  const hashHex = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('')
  return hashHex
}

/**
 * Checks if the email is a Postguard mail based on the attachment content type
 * @return True or false
 */
export function isPostGuardEmail(): boolean {
  if (Office.context.mailbox.item.attachments.length != 0) {
    const attachmentContentType =
      Office.context.mailbox.item.attachments[0].contentType
    if (attachmentContentType == 'application/postguard') {
      utilLog.info('It is a PostGuard email!')
      return true
    } else {
      utilLog.info('No PostGuard email')
      return false
    }
  } else {
    utilLog.info('No PostGuard email')
    return false
  }
}

/**
 * Creates a readable stream from an array
 * @param array The array to be converted into a stream
 * @returns The readable stream
 */
export function newReadableStreamFromArray(array) {
  return new ReadableStream({
    start(controller) {
      controller.enqueue(array)
      controller.close()
    }
  })
}

export function getGlobal() {
  return typeof self !== 'undefined'
    ? self
    : typeof window !== 'undefined'
    ? window
    : typeof global !== 'undefined'
    ? // eslint-disable-next-line no-undef
      global
    : undefined
}

export function getPostGuardHeaders() {
  // See https://learn.microsoft.com/en-us/javascript/api/office/office.platformtype
  const host =
    Office.context.platform.toString() === 'OfficeOnline'
      ? 'OutlookWeb'
      : ['iOS', 'Android'].includes(Office.context.platform.toString())
      ? 'OoutlookMobile'
      : 'OutlookDesktop'

  let headers = `${host},${Office.context.diagnostics.version},pg4ol,0.0.1`
  console.log(`Headers: ${headers}`)
  return headers
}
