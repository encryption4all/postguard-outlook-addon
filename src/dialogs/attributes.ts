/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global $, Office */

// images references in the manifest
import { Policy, getGlobal } from '../helpers/utils'
import { AttributeForm } from '@e4a/pg-components'

import '../../assets/16.png'
import '../../assets/32.png'
import '../../assets/80.png'

// eslint-disable-next-line no-undef
const getLogger = require('webpack-log')
const attributeLog = getLogger({ name: 'PostGuard attribute log' })

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

const g = getGlobal() as any

/**
 * Initialization function which also extracts the URL params
 */
Office.initialize = function () {
  if (Office.context.mailbox === undefined) {
    attributeLog.info('Add attributes dialog openend!')
    const urlParams = new URLSearchParams(window.location.search)
    g.token = Buffer.from(urlParams.get('token'), 'base64').toString('utf-8')
    const senderObj = urlParams.get('sender')

    let recipients
    let sender

    $(function () {
      const el = document.querySelector('#root')
      if (!el) return

      if (senderObj !== null) {
        sender = Buffer.from(senderObj, 'base64').toString('utf-8')
        g.signDialog = true
      } else {
        recipients = JSON.parse(
          Buffer.from(urlParams.get('recipients'), 'base64').toString('utf-8')
        )
      }

      const start = senderObj !== null ? [sender] : recipients

      const init = start.reduce((policies, next) => {
        const email = next
        policies[email] = []
        return policies
      }, [])

      const customText = senderObj !== null ? 'Next' : 'Send'

      new AttributeForm({
        target: el,
        props: {
          initialPolicy: init,
          signing: senderObj !== null,
          onSubmit: finish,
          submitButton: { customText: customText },
          lang: g.language
        }
      })

      attributeLog.info(
        `Token: ${g.token}, recipients: ${recipients}, sender: ${sender}`
      )
    })
  }
}

const finish = async (policy: Policy) => {
  const msg = {
    result: { policy: policy, accessToken: g.token },
    operation: g.signDialog === true ? 'sign' : 'enc',
    status: 'success'
  }
  passMsgToParent(JSON.stringify(msg))
}

/**
 * Passes a message to the parent
 * @param msg The message
 */
function passMsgToParent(msg: string) {
  if (Office.context.mailbox === undefined) {
    Office.context.ui.messageParent(msg)
  }
}
