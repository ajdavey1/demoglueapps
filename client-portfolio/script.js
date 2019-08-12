window.applicationName = 'Vanilla Client Portfolio'
import {onSyncContact, onExcelStatusChange, onOutlookStatusChange, announcePortfolioChange, onPortfolioUpdateAnnounce, setTitle, isUsingChannels} from './glue-related.js'
import {getRestId, getIntialClientId, setButtonAvailability} from '../shared/utils.js';
import {sendEmail} from '../shared/send-email.js';
import {openSheet} from '../shared/open-sheet.js'

let displayedContact = undefined;
let acceptSync = false;
let sheetSubscription = false;

(async function init() {
  addClickListener();
  trackOfficeAddins();

  let usingChannels =  await isUsingChannels();

  if (getIntialClientId() && !usingChannels) {
    loadContact(getIntialClientId());
  } else {
    changeSync(true)
  }

  onSyncContact((contact) => {
    if (acceptSync) {
      displayContact(contact);
      displayedContact = contact;

      if (sheetSubscription) {
        sheetSubscription();
        sheetSubscription = false;
      }
    }
  });

  onPortfolioUpdateAnnounce(({clientId, portfolio}) => {
    if (getRestId(displayedContact) === clientId) {
      displayedContact.context.portfolio = portfolio;
      displayContact(displayedContact);
    }
  })
}())

function trackOfficeAddins() {
  let openSheetBtn = document.querySelector('[action="open-sheet"]');
  let sendEmailBtn = document.querySelector('[action="send-email"]');

  onExcelStatusChange((connected) => setButtonAvailability(openSheetBtn, connected))
  onOutlookStatusChange((connected) => setButtonAvailability(sendEmailBtn, connected))
}

function addClickListener() {
  document.addEventListener('click', (event) => {
    if (event.target.matches('[action], [action] *')) {
      let button = event.path.reduce((acc, cur) => {
        return cur.matches && cur.matches('[action]') ? cur : acc;
      });

      if (button.hasAttribute('disabled')) {
        return;
      }

      let action = button.getAttribute('action');
      switch(action) {
        case 'send-email': emailContact(); break;
        case 'open-sheet': openSheetForContact(); break;
        case 'sync-on': changeSync(true); break;
        case 'sync-off': changeSync(false); break;
      }
    }
  })
}

async function loadContact(contactId) {
  let contact = await (await fetch(`http://localhost:22060/clients/${contactId}`)).json();
  console.log(contact);
  if (!contact) {
    return;
  }

  displayContact(contact);
  displayedContact = contact;
}

function displayContact(contact) {
  clearTable();
  if (!contact) {
    return;
  }

  document.querySelector('#title-name').innerHTML = contact.displayName;

  let emptyRow = document.querySelector('.empty-row').cloneNode(true);
  emptyRow.style.display = '';
  emptyRow.classList.remove('empty-row')

  contact.context.portfolio.forEach(instrument => {
    let newRow = emptyRow.cloneNode(true);
    newRow.querySelector('.instrument-ric').innerText = instrument.ric;
    newRow.querySelector('.instrument-name').innerText = instrument.description;
    newRow.querySelector('.instrument-price').innerText = instrument.price;
    newRow.querySelector('.instrument-number').innerText = instrument.shares;

    document.querySelector('table tbody').appendChild(newRow);
  });

  setTitle(contact.displayName)
}

function emailContact() {
  sendEmail(displayedContact)
}

async function openSheetForContact() {
  sheetSubscription = await openSheet(displayedContact, (data) => {
    displayedContact.context.portfolio = data;
    displayContact(displayedContact);
    updateContactToServer(displayedContact);
    announcePortfolioChange(displayedContact, data)
  });
}

function updateContactToServer(contact) {
  let url = `http://localhost:22060/clients/${getRestId(contact)}`;
  fetch(url,  {
    method: 'PUT',
    body: JSON.stringify(contact),
    headers:{
      'Content-Type': 'application/json'
    }
  }).then(console.log)
}

function changeSync(newValue) {
  acceptSync = newValue;

  document.querySelector(`[action="sync-${newValue ? 'on' : 'off'}"]`).classList.add('hidden');
  document.querySelector(`[action="sync-${newValue ? 'off' : 'on'}"]`).classList.remove('hidden');
}

function clearTable() {
  document.querySelectorAll('tbody tr:not(.empty-row)').forEach(row => {
    row.remove();
  })
}