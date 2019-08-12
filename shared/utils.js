/* global Glue */

function getPortfolioValue(portfolio) {
  let portfolioValue = portfolio.reduce((sum, ticker) => {
    sum += ticker.price * ticker.shares;
    return sum;
  }, 0);

  return portfolioValue.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

function getRestId(contact) {
  let ids;
  let id;
  if (contact.ids) {
    ids = contact.ids;
  } else if(Array.isArray(contact)) {
    ids = contact;
  }

  ids.forEach(currentId => {
    if (currentId.systemName === 'rest.id') {
      id = currentId.nativeId;
    }
  });

  return id;
}

function cloneObject(object) {
  return JSON.parse(JSON.stringify(object));
}

function getIntialClientId() {
  let queryParmsClientId = getQueryParams()['clientId'];
  let contextClientId = window.glue42gd && glue42gd.context.clientId;

  return queryParmsClientId || contextClientId;
}

function getQueryParams() {
  return location.search
    .slice(1)
    .split('&')
    .filter(v => v)
    .reduce((acc, cur) => {
    let [key, value] = cur.split('=');
    acc[key] = value;
    return acc;
  }, {})
}

function setButtonAvailability(button, status) {
  if (status) {
    button.removeAttribute('disabled');
    button.classList.remove('disabled');
  } else {
    button.setAttribute('disabled', true);
    button.classList.add('disabled');
  }
}

export {
  getPortfolioValue,
  getRestId,
  cloneObject,
  getQueryParams,
  getIntialClientId,
  setButtonAvailability
}