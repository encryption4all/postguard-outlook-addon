/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, Office */

import 'web-streams-polyfill'

import {
  getPostGuardHeaders,
  PKG_URL,
  _getMobileUrl,
  getGlobal
} from '../helpers/utils'

import * as YiviCore from '@privacybydesign/yivi-core'
import * as YiviClient from '@privacybydesign/yivi-client'
import * as YiviWeb from '@privacybydesign/yivi-web'

import * as getLogger from 'webpack-log'
import i18next from 'i18next'
import { AttributeCon } from '@e4a/pg-wasm'
const encryptLog = getLogger({ name: 'PostGuard encrypt log' })

// eslint-disable-next-line no-undef
var Buffer = require('buffer/').Buffer

const g = getGlobal() as any

/**
 * Initialization function which also extracts the URL params
 */
Office.initialize = function () {
  if (Office.context.mailbox === undefined) {
    encryptLog.info('Sign message dialog openend!')
    const urlParams = new URLSearchParams(window.location.search)
    g.token = Buffer.from(urlParams.get('token'), 'base64').toString('utf-8')
    const policyArg = Buffer.from(
      urlParams.get('signingpolicy'),
      'base64'
    ).toString('utf-8')
    const policy: AttributeCon = JSON.parse(policyArg)

    $(function () {
      executeIrmaDisclosureSession(policy)
    })
  }
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

/**
 * Executes an IRMA disclosure session based on the policy for signing
 * @param policy The policy
 * @returns The JWT of the IRMA session
 */
async function executeIrmaDisclosureSession(policy: AttributeCon) {
  // show HTML elements needed
  document.getElementById('header_text').style.display = 'block'
  document.getElementById('decryptinfo').style.display = 'block'
  document.getElementById('irmaapp').style.display = 'block'
  document.getElementById('qrcodecontainer').style.display = 'block'
  document.getElementById('loading').style.display = 'none'

  // calculate diff in seconds between now and tomorrow 4 am
  let tomorrow = new Date()
  tomorrow.setDate(tomorrow.getDate() + 1)
  tomorrow.setHours(4, 0, 0, 0)
  const now = new Date()
  const seconds =
    Math.floor((tomorrow.getTime() - now.getTime()) / 1000) % 86400
  encryptLog.info('Diff in seconds until 4 am: ', seconds)

  const requestBody = {
    con: policy,
    validity: seconds
  }

  const yivi = new YiviCore({
    debugging: true,
    element: '#qrcode',
    language: 'en',
    state: {
      serverSentEvents: false,
      polling: {
        endpoint: 'status',
        interval: 500,
        startState: 'INITIALIZED'
      }
    },
    session: {
      url: PKG_URL,
      start: {
        url: (o) => `${o.url}/v2/request/start`,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Postguard-Client-Version': getPostGuardHeaders()
        },
        body: JSON.stringify(requestBody)
      },
      result: {
        url: (o, { sessionToken: token }) => `${o.url}/v2/request/jwt/${token}`,
        parseResponse: (r) => {
          return new Promise((resolve, reject) => {
            if (r.status != '200') reject('not ok')
            resolve(r.url)
          })
        }
      }
    }
  })
  yivi.use(YiviClient)
  yivi.use(YiviWeb)
  // disclose and retrieve JWT URL
  const jwtUrl = await yivi.start()
  const localJwt: string = await $.ajax({ url: jwtUrl })

  const msg = {
    result: { jwt: localJwt, accessToken: g.token },
    status: 'success'
  }
  passMsgToParent(JSON.stringify(msg))
}
