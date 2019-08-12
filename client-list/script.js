import { onSyncContact, syncContact, contactActions, onPortfolioUpdateAnnounce,onExcelStatusChange, onOutlookStatusChange } from './glue-related.js';
import { getPortfolioValue, getRestId, cloneObject } from '../shared/utils.js';
import { sendEmail } from '../shared/send-email.js';

document.addEventListener('DOMContentLoaded', async () => {
  displayContacts();
  addSearchFilter();
  addCallbacks();
  addClickListener();
  trackAddins();
})

window.dc = displayContacts;

async function displayContacts() {
  let contacts = (window.contacts || await getContacts());
  window.contacts = contacts;

  let emptyRow = document.querySelector('.empty-row').cloneNode(true);
  emptyRow.style.display = '';
  emptyRow.classList.remove('empty-row');

  // document.querySelector('.contacts-table tbody tr').innerHTML = '';
  document.querySelectorAll('.contacts-table tbody tr:not(.empty-row)').forEach(n => n.remove());

  contacts.forEach(contact => {
    // clone the empty "template" row, fill current contact details and add it to the table
    const {displayName, emails, context:{portfolio}, _id} = contact;
    const newRow = emptyRow.cloneNode(true);
    let portfolioValue = getPortfolioValue(portfolio);

    newRow.querySelector('.client-name').innerText = displayName;
    newRow.querySelector('.client-email').innerText = emails[0];
    newRow.querySelector('.client-portfolio').innerText = portfolioValue;
    newRow.setAttribute('client-id', _id)


    document.querySelector('.contacts-table tbody').appendChild(newRow);
  });
}

async function getContacts() {
  console.log('fetching contacts');
  const restURL = 'http://localhost:22060/clients';
  return (await fetch(restURL)).json();
}

function addSearchFilter() {
  const filterInput = document.querySelector('input.contacts-filter');
  filterInput.addEventListener('keyup', () => {
    const filterValue = filterInput.value.toLowerCase();

    document.querySelectorAll('.contacts-table tbody tr:not(.empty-row)').forEach(row => {
      const currentContactName = row.querySelector('.client-name').innerText.toLowerCase();
      const currentContactEmail = row.querySelector('.client-email').innerText.toLowerCase();
      const matches = (currentContactName.indexOf(filterValue) >= 0) || (currentContactEmail.indexOf(filterValue) >= 0);
      // hide or show the current row depending if the search value matches the name or the email
      row.style.display = matches ? '' : 'none';
    })
  })
}

async function addCallbacks() {
  onSyncContact((contact) => {
    clearRowsSelection();
    let restId = getRestId(contact);
    selectRow(restId);
  });

  onPortfolioUpdateAnnounce(({clientId, portfolio}) => {
    contacts.forEach(contact => {
      if (getRestId(contact) === clientId) {
        // console.log();
        contact.context.portfolio = portfolio;
        displayContacts();
      }
    })
  })
}

function selectRow(restId) {
  document.querySelectorAll(`.contacts-table tbody tr[client-id].bg-primary`).forEach(el => el.classList.remove('bg-primary'))

  let selectedContactRow = document.querySelector(`.contacts-table tbody tr[client-id="${restId}"]`);
  if (selectedContactRow) {
    selectedContactRow.classList.add('bg-primary');
  }

  // let table = document.querySelector('.contacts-table');
  // let tableTop = table.getBoundingClientRect().top;
  // let tableBottom = table.getBoundingClientRect().bottom;
  // let rowTop = selectedContactRow.getBoundingClientRect().top;
  // if ((rowTop < tableTop) || (rowTop > tableBottom)) {
  //   document.querySelector('html').scrollTo({
  //     left: 0,
  //     top: rowTop - tableTop,
  //     behavior: 'smooth'
  //   })
  // }
}

function clearRowsSelection() {
  document.querySelectorAll('.contacts-table tbody tr.bg-primary').forEach(node => {
    node.classList.remove('bg-primary');
  });
}

function getContactById(restId) {
  return contacts.reduce((acc, currentContact) => {
    if (getRestId(currentContact) === restId) {
      return currentContact
    } else {
      return acc;
    }
  }, null);
}

function syncContactById(restId) {
  const contact = getContactById(restId);
  syncContact(contact);
}

function addClickListener() {
  document.addEventListener('click', (event) => {
    handleMenuClick(event);

    if (event.target.matches('.dropdown-item[action]')) {
      handleActionClick(event);
    }

    if (event.target.matches('.contacts-table tr, .contacts-table tr *') && !event.target.matches('.contacts-table tr td:first-child, .contacts-table tr td:first-child *')) {
      handleRowClick(event);
    }
  })
}

function handleMenuClick(event) {
  if (event.target.matches('.contacts-table tbody tr td.action-button .dropdown, .contacts-table tbody tr td.action-button .dropdown *')) {
    event.path.forEach(element => {
      if (element.matches && element.matches('tr')) {
        let menu = element.querySelector('.dropdown-menu');
        let menuIsVisible = menu.classList.contains('show');
        closeAllMenus();
        if (!menuIsVisible) {
          menu.classList.add('show');
        }
      }
    })
  } else {
    closeAllMenus();
  }
}

function handleActionClick(event) {
  event.path.forEach(element => {
    if (element.matches && element.matches('tr')) {
      let restId = element.getAttribute('client-id');
      let contact = getContactById(restId);
      let action  = event.target.getAttribute('action');

      switch(action) {
        case 'openPortfolio': {
          contactActions.openPortfolio(restId, contact);
          break;
        }
        case 'openPortfolioInExcel': {
          contactActions.openPortfolioInExcel(restId, contact, onSheetChanged);
          break;
        }
        case 'openContact': {
          contactActions.openContact(restId, contact);
          break;
        }
        case 'emailContact': {
          contactActions.emailContact(restId, contact);
          break;
        }
        case 'updateContact': {
          contactActions.updateContact(restId, contact);
          break;
        }
      }
    }
  });
}

function onSheetChanged(changedContact, newPortfolio) {
  contacts.forEach(contact => {
    if (getRestId(contact) === getRestId(changedContact)) {
      contact.context.portfolio = newPortfolio;
      updateContactToServer(contact);
      displayContacts();
    }
  })
}

function handleRowClick(event) {
  event.path.forEach(element => {
    if (element.matches && element.matches('tr')) {
      let clientId = element.getAttribute('client-id');
      syncContactById(clientId);
      selectRow(clientId);
    }
  });
}

function closeAllMenus() {
  document.querySelectorAll('.dropdown-menu.show').forEach(menu => {
    menu.classList.remove('show');
  })
}

function trackAddins() {
  onExcelStatusChange((newStatus) => {
    setButtonsState(document.querySelectorAll('[action="openPortfolioInExcel"]'), newStatus);
  });

  onOutlookStatusChange((newStatus) => {
    setButtonsState(document.querySelectorAll('[action="emailContact"], [action="updateContact"]'), newStatus);
  })
}

function setButtonsState(buttons, status) {
  buttons.forEach(button => {
    if (status) {
      button.classList.remove('disabled');
    } else {
      button.classList.add('disabled');
    }
  })
}

function updateContactToServer(contact) {
  const restId = getRestId(contact);

  fetch(`http://localhost:22060/clients/${restId}/portfolio`, {
    method: 'PUT',
    body: JSON.stringify(contact.context.portfolio),
    headers:{
      'Content-Type': 'application/json'
    }
  }).then((response) => {
    response.text()
      .then(console.log)
  })
}