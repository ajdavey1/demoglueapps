const glueReadyPromise = new Promise(async (resolve, reject) => {
  let glueConfig = await getGlueConfig();

  console.log(glueConfig);

  Glue(glueConfig)
    .then(glue => {
      Glue4Office({glue, excel: true, outlook: true, word: false})
        .then(g4o => {
          window.glue = g4o;
          resolve(g4o)
        })
    }).catch((err) => {

      console.log(err);
    })
});

async function getGlueConfig() {
  let auth;
  let appManager = true;
  let inContainer = !!window.glue42gd;

  if (inContainer) {
    auth = {gatewayToken: await glue42gd.getGWToken()}
    appManager = 'full';
  } else {
    auth = {username: 'aatanasov', password: 't42demo'};
  }

  let gwUrl = inContainer ? glue42gd.gwURL : 'ws://localhost:8385';

  let config = {
    appManager,
    gateway: {
      ws: gwUrl,
      protocolVersion: 3
    },
    auth: auth,
    channels: true
  };

  return config;
}

export {glueReadyPromise};