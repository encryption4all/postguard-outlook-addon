/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global $, Office */

const defaultValues = new Map([
  ['store-unencrypted', 'yes'],
  ['subject-encrypted', 'no']
])

let settingsSavedAmount = 0
//Array.from(defaultValues.keys()).length

;(function () {
  'use strict'

  // The Office initialize function must be run each time a new page is loaded.
  // eslint-disable-next-line no-unused-vars
  Office.initialize = function (_init) {
    jQuery(function () {
      Array.from(defaultValues.keys()).forEach((key) => loadRoamingSetting(key))

      $('#settings-done').on('click', function () {
        Array.from(defaultValues.keys()).forEach((key) =>
          saveToRoamingSetting(key)
        )
      })
    })
  }

  // Saves all roaming settings.
  function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.log('Error during saving settings')
    } else {
      settingsSavedAmount++
    }
  }

  // Save setting to roaming settings
  function saveToRoamingSetting(name: string) {
    const _settings = Office.context.roamingSettings
    const selectedVal = $(`input[name=${name}]:checked`)
      .attr('id')
      .endsWith('yes')
      ? 'yes'
      : 'no'
    _settings.set(name, selectedVal)
    _settings.saveAsync(saveMyAppSettingsCallback)
    console.log(`Save ${name} with value:`, _settings.get(name))
  }

  // Load stored setting or use default value
  function loadRoamingSetting(name: string) {
    const _settings = Office.context.roamingSettings
    let settingValue = _settings.get(name)
    if (settingValue === undefined) {
      settingValue = defaultValues.get(name)
    }
    console.log(`Setting ${name} has value ${settingValue}`)

    $(`#${name}-${settingValue}`).attr('checked', 'checked')
  }
})()
