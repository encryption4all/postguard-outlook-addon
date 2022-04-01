/* global $, Office */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'

export const htmlBodyType: MicrosoftGraph.BodyType = 'html'

export interface IAttachmentContent {
  filename: string
  content: string
  isInline: boolean
}

// 1. replace mail body
// 2. remove current attachment
// 3. add attachments
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
      console.log('PATCH message success: ', success)
      removeAttachment(token, itemId, item.attachments[0].id, attachments)
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

function removeAttachment(
  token: string,
  itemId: string,
  attachmentId: string,
  attachments: IAttachmentContent[]
) {
  const attachmentUrl = `https://graph.microsoft.com/v1.0/me/messages/${itemId}/attachments/${attachmentId}`
  $.ajax({
    type: 'DELETE',
    url: attachmentUrl,
    headers: {
      Authorization: 'Bearer ' + token
    },
    success: function (success) {
      console.log('DELETE attachment success: ', success)
      attachments.forEach((attachment) => {
        addAttachment(token, itemId, attachment)
      })
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

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
      console.log('Add attachment success: ', success)
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

// get id from PostGuard folder to create inner mail in that folder
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
        createPostGuardMailFolder(token, innerMail, attachments, folder_name)
      }
    }
  }).fail(function ($xhr) {
    var data = $xhr.responseJSON
    console.log('Ajax error: ', data)
    setEventError()
  })
}

function createPostGuardMailFolder(token, innerMail, attachments, folder_name) {
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

var mailPopup

export function showMailPopup(mailBody: string) {
  var fullUrl =
    location.protocol +
    '//' +
    location.hostname +
    (location.port ? ':' + location.port : '') +
    '/emailpopup.html'

  window.localStorage.setItem('mailBody', mailBody)
  Office.context.ui.displayDialogAsync(
    fullUrl,
    { height: 60, width: 60 },
    function (asyncResult) {
      mailPopup = asyncResult.value
      mailPopup.messageChild(mailBody)
    }
  )
}

export function utilsFillMailPopup(content: string) {
  mailPopup.messageChild(content)
}

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

export async function hashString(message: string): Promise<string> {
  const msgArray = new TextEncoder().encode(message)
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgArray)
  const hashArray = Array.from(new Uint8Array(hashBuffer))
  const hashHex = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('')
  return hashHex
}
