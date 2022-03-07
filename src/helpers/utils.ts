/* global $, Office */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'

export const htmlBodyType: MicrosoftGraph.BodyType = 'html'

export interface IAttachmentContent {
  filename: string
  content: string
  isInline: boolean
}

// get id from cryptify folder to create inner mail in that folder
// if it does not exist, create it
export function storeMailAsPlainLocally(
  token: string,
  innerMail: MicrosoftGraph.Message,
  attachments: IAttachmentContent[],
  folder_name: string
) {
  const mailFoldersUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders'
  $.ajax({
    type: 'GET',
    url: mailFoldersUrl,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      console.log('MailFolders: ', success)
      let folderFound = false
      success.value.forEach((folder) => {
        if (!folderFound && folder.displayName === folder_name) {
          folderFound = true
          console.log('Folder exists with id ', folder.id)
          storeInnerMail(folder.id, innerMail, token, attachments)
        }
      })
      if (!folderFound) {
        console.log('Folder not found, creating ...')
        createCryptifyMailFolder(token, innerMail, attachments, folder_name)
      }
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

function createCryptifyMailFolder(token, innerMail, attachments, folder_name) {
  const createMailFoldersUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders'
  const payload = {
    displayName: folder_name,
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
      console.log('Created mailfolder succesfully!')
      storeInnerMail(success.id, innerMail, token, attachments)
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

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
      console.log('Createmail success: ', success)
      attachments.forEach((attachment) => {
        storeAttachment(folderId, success.id, token, attachment)
      })
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

async function storeAttachment(
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
      console.log('Create attachment success: ', success)
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

export function setEventError() {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: 'Error during encryption, please contact your administrator.'
  }

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    'action',
    message
  )
}
