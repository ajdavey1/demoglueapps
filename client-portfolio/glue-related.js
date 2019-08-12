import {glueReadyPromise} from '../shared/initialize-glue.js';
import { getRestId } from '../shared/utils.js';

glueReadyPromise
  .then((glue) => {
    registerMethods();
    trackAddins();
  })

async function registerMethods() {
  let glue = await glueReadyPromise;

  let usingChannels = window.glue42gd && glue.appManager.application(glue42gd.applicationName).userProperties.useChannels;

  if (!usingChannels) {
    glue.agm.register('T42.CRM.SyncContact', (data, server) => {
      // console.log(data, server);
      if (server.instance === glue.agm.instance.instance) { return; }

      onSyncContactCallbacks.forEach(callback => {
        callback(data.contact);
      });
    });
  } else {
    glue.channels.subscribe((payload, meta) => {
      onSyncContactCallbacks.forEach(callback => {
        callback(payload.contact);
      });
    });
  }

  glue.agm.register('UpdatePortfolio', (data, server) => {
    updatePortfolioCallbacks.forEach(callback => {
      callback(data);
    })
  })
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

let onSyncContactCallbacks = [];
function onSyncContact(callback) {
  onSyncContactCallbacks.push(callback)
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

let onSheetChangedCallbacks = [];
function onSheetChanged(callback) {
  onSheetChangedCallbacks.push(callback);
}

let updatePortfolioCallbacks = [];
function onPortfolioUpdateAnnounce(callback) {
  updatePortfolioCallbacks.push(callback);
}

async function announcePortfolioChange(contact, data) {
  let restId = getRestId(contact);
  let glue = await glueReadyPromise;
  glue.agm.invoke('UpdatePortfolio', {
    clientId: restId,
    portfolio: data,
    errors: []
  }, 'all')
}

async function setTitle(newTitle) {
  document.title = newTitle;
  let glue = await glueReadyPromise;
  if (glue && glue.windows && glue.windows.my) {
    glue.windows.my().setTitle(newTitle)
  }
}

async function isUsingChannels() {
  let glue = await glueReadyPromise;
  return glue.appManager.myInstance.application.userProperties.useChannels;
}


export {
  onSyncContact,
  onExcelStatusChange,
  onOutlookStatusChange,
  announcePortfolioChange,
  onPortfolioUpdateAnnounce,
  setTitle,
  isUsingChannels
}