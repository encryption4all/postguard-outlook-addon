/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global $, Office */

// images references in the manifest
import { getGlobal } from '../helpers/utils'
import '../../assets/16.png'
import '../../assets/32.png'
import '../../assets/80.png'

import AttributeForm from 'attribute-form/AttributeForm/AttributeForm.svelte'
import type { Policy } from 'attribute-form/AttributeForm/AttributeForm.svelte'

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
    const recipients = JSON.parse(
      Buffer.from(urlParams.get('recipients'), 'base64').toString('utf-8')
    )

    attributeLog.info(`Token: ${g.token}, recipients: ${g.recipients}`)

    $(function () {
      const el = document.querySelector('#root')
      if (!el) return

      const init = recipients.reduce((policies, next) => {
        const email = next
        policies[email] = []
        return policies
      }, [])

      new AttributeForm({
        target: el,
        props: {
          initialPolicy: init,
          onSubmit: finish,
          submitButton: { customText: 'Send' }
        }
      })
    })
  }
}

const finish = async (policy: Policy) => {
  const msg = {
    result: { policy: policy, accessToken: g.token },
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
