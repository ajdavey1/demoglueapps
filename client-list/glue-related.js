import {
  glueReadyPromise
} from '../shared/initialize-glue.js';

import {
  sendEmail
} from '../shared/send-email.js';
import {
  openSheet
} from '../shared/open-sheet.js';

let usingChannels = false;

glueReadyPromise.then(glue => {
  usingChannels = window.glue42gd && glue.appManager.application(glue42gd.applicationName).userProperties.useChannels;
  if (!usingChannels) {
    glue.agm.register('T42.CRM.SyncContact', (payload, server) => {
      if (server.instance === glue.agm.instance.instance) {
        // this is a call from us
        return;
      }

      onSyncCallbacks.forEach(callback => {
        callback(payload.contact)
      })
    })
  } else {
    glue.channels.subscribe((payload, meta) => {
      onSyncCallbacks.forEach(callback => {
        callback(payload.contact)
      })
    })
  }

  glue.agm.register('UpdatePortfolio', (data, server) => {
    updatePortfolioCallbacks.forEach(callback => {
      callback(data);
    })
  })

  trackAddins();
})

let onSyncCallbacks = [];

function onSyncContact(callback) {
  onSyncCallbacks.push(callback);
}

function trackAddins() {
  excelStatusChangeCallbacks.forEach(callback => {
    callback(glue.excel.addinStatus)
  });

  outlookStatusChangeCallbacks.forEach(callback => {
    callback(glue.outlook.addinStatus)
  })

  glue.excel.onAddinStatusChanged((connected) => {
    excelStatusChangeCallbacks.forEach(callback => {
      callback(connected)
    });
  });

  glue.outlook.onAddinStatusChanged(({connected}) => {
    outlookStatusChangeCallbacks.forEach(callback => {
      callback(connected);
    })
  });
}

async function syncContact(contact) {
  await glueReadyPromise;
  if (usingChannels) {
    glue.channels.publish({
      contact
    })
  } else {
    return glue.agm.invoke('T42.CRM.SyncContact', {
      contact
    }, 'all');
  }
}

async function openPortfolio(restId, contact) {
  if (window.glue42gd) {
    await glueReadyPromise;
    let portfolioAppName = (usingChannels ? 'channels' : '') + 'vanillaclientportfolio'
    glue.appManager.application(portfolioAppName)
      .start({
        clientId: restId
      }, {
        relativeDirection: 'right',
        relativeTo: glue42gd.windowId,
        tabGroupId: glue42gd.windowId
      })
  } else {
    window.open(`../client-portfolio?clientId=${restId}`);
  }
}

function openPortfolioInExcel(restId, contact, callback) {
  return openSheet(contact, (data) => {
    callback(contact, data);
    announcePortfolioChange(restId, data)
  })
}

async function openContact(restId, contact) {
  if (window.glue42gd) {
    await glueReadyPromise;
    let contactAppName = (usingChannels ? 'channels' : '') + 'clientcontact'
    glue.appManager.application(contactAppName)
      .start({
        clientId: restId
      }, {
        relativeDirection: 'right',
        relativeTo: glue42gd.windowId,
        tabGroupId: glue42gd.windowId
      })
  } else {
    window.open(`http://localhost:22080/client-list-portfolio-contact/dist/#/clientcontact/${restId}`);
  }
}

function emailContact(restId, contact) {
  sendEmail(contact);
}

function updateContact(restId, contact) {
  glueReadyPromise.then(glue => {
    console.log('update contact ', restId);
  })
}

async function announcePortfolioChange(restId, data) {
  let glue = await glueReadyPromise;
  glue.agm.invoke('UpdatePortfolio', {
    clientId: restId,
    portfolio: data,
    errors: []
  }, 'all')
}

let updatePortfolioCallbacks = [];

function onPortfolioUpdateAnnounce(callback) {
  updatePortfolioCallbacks.push(callback);
}

function isUsingChannels() {

}

let excelStatusChangeCallbacks = [];
async function onExcelStatusChange(callback) {
  excelStatusChangeCallbacks.push(callback);
  let glue = await glueReadyPromise;
  callback(glue.excel.addinStatus);
}

let outlookStatusChangeCallbacks = [];
async function onOutlookStatusChange(callback) {
  outlookStatusChangeCallbacks.push(callback);
  let glue = await glueReadyPromise;
  callback(glue.outlook.addinStatus);
}



const contactActions = {
  openPortfolio,
  openPortfolioInExcel,
  openContact,
  emailContact,
  updateContact
}

export {
  glueReadyPromise,
  onSyncContact,
  syncContact,
  contactActions,
  onPortfolioUpdateAnnounce,
  onExcelStatusChange,
  onOutlookStatusChange
}