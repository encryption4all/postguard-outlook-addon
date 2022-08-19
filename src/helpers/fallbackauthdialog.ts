/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

/* global Office */

import {
  AccountInfo,
  Configuration,
  PublicClientApplication,
  RedirectRequest
} from '@azure/msal-browser'

const msalConfig: Configuration = {
  auth: {
    clientId: '6ee2a054-1d61-405d-8e5d-c2daf25c5833'
  },
  cache: {
    cacheLocation: 'localStorage', // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true // Recommended to avoid certain IE/Edge issues.
  }
}

const requestObj: RedirectRequest = {
  scopes: [
    'Mail.ReadBasic',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'openid',
    'profile',
    'offline_access'
  ]
}

const publicClientApp: PublicClientApplication = new PublicClientApplication(
  msalConfig
)

Office.onReady(async () => {
  if (Office.context.ui.messageParent) {
    try {
      let tokenResponse = await publicClientApp.handleRedirectPromise()

      let accountObj: AccountInfo
      if (tokenResponse) {
        accountObj = tokenResponse.account
      } else {
        const urlParams = new URLSearchParams(window.location.search)
        const currentAccountMail = urlParams.get('currentAccountMail')
        localStorage.setItem('authCurrentMail', currentAccountMail)

        const allAccountsurrentAccount: AccountInfo[] =
          publicClientApp.getAllAccounts()
        accountObj = allAccountsurrentAccount.find((acc) =>
          findMyAcc(acc, currentAccountMail)
        )
        publicClientApp.setActiveAccount(accountObj)
      }

      if (accountObj && tokenResponse) {
        setAuthLog('[AuthService.init] Got valid accountObj and tokenResponse')
        handleResponse(tokenResponse, 'Got valid accountObj and tokenResponse')
      } else if (accountObj) {
        setAuthLog('[AuthService.init] User has logged in, but no tokens.')
        try {
          tokenResponse = await publicClientApp.acquireTokenSilent({
            account: accountObj,
            scopes: requestObj.scopes
          })
          handleResponse(tokenResponse, 'User has logged in, but no tokens.')
        } catch (err) {
          await publicClientApp.acquireTokenRedirect(requestObj)
        }
      } else {
        setAuthLog(
          '[AuthService.init] No accountObject or tokenResponse present. User must now login.'
        )
        await publicClientApp.loginRedirect(requestObj)
      }
    } catch (error) {
      setAuthLog(
        '[AuthService.init] Failed to handleRedirectPromise(): ' +
          JSON.stringify(error)
      )
    }
  }
})

function setAuthLog(message: string) {
  localStorage.setItem(
    'authLog',
    new Date() + ': ' + message + ', ' + localStorage.getItem('authLog')
  )
}

function findMyAcc(user: AccountInfo, currentAccountMail: string) {
  return user.username.toLowerCase() === currentAccountMail.toLowerCase()
}

function handleResponse(response, message = null) {
  if (response.tokenType === 'id_token') {
    localStorage.setItem('loggedIn', 'yes')
  } else {
    console.log('token type is:' + response.tokenType)
    Office.context.ui.messageParent(
      JSON.stringify({ status: 'success', result: response, message: message })
    )
  }
}

/*localStorage.setItem('activeAccount', JSON.stringify(currentAccount))

    localStorage.setItem(
      'authAccounts',
      JSON.stringify(allAccountsurrentAccount)
    )

    const silentFlowRequest: SilentRequest = {
      scopes: requestObj.scopes,
      account: currentAccount
    }

    publicClientApp
      .acquireTokenSilent(silentFlowRequest)
      .then(function (accessTokenResponse) {
        handleResponse(accessTokenResponse, 'silent successfull')
      })
      // eslint-disable-next-line no-unused-vars
      .catch(function (error) {
        localStorage.setItem('authError', JSON.stringify(error))
        if (response) {
          handleResponse(response, 'handle redirect response')
        }
        // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
        // stored login data in localStorage. So a direct call of acquireTokenRedirect
        // causes the error "User login is required". Once the user is logged in successfully
        // the first time, msal data in localStorage will prevent this error from ever hap-
        // pening again; but the error must be blocked here, so that the user can login
        // successfully the first time. To do that, call loginRedirect first instead of
        // acquireTokenRedirect.
        else if (localStorage.getItem('loggedIn') === 'yes') {
          publicClientApp.acquireTokenRedirect(requestObj)
        } else {
          // This will login the user and then the (response.tokenType === "id_token")
          // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
          // and then the dialog is redirected back to this script, so the
          // acquireTokenRedirect above runs.
          publicClientApp.loginRedirect(requestObj)
        }
      })
  }
})
*/
