function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const token = normalizeSessionToken_(params.token);
  let page = params.page || 'Dashboard';
  const id = params.id || '';

  page = String(page).replace(/^Views\//, '');
  const session = token ? getValidSession(token) : null;

  if (!session) {
    const template = HtmlService.createTemplateFromFile('Views/Login');
    template.scriptUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle('Investigasi Pertambangan Batubara — Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // google.script.run is a separate server execution from doGet().
  // Store the token here so RPC calls can recover the same validated session.
  registerSessionBridge_(session.Token);

  let template;
  try {
    template = HtmlService.createTemplateFromFile('Views/' + page);
  } catch (err) {
    template = HtmlService.createTemplateFromFile('Views/Dashboard');
    page = 'Dashboard';
  }

  template.userName = session.Nama || session.Username;
  template.userRole = session.Role;
  template.userSite = session.Site;
  template.token = session.Token || token;
  template.idInsiden = id;
  template.page = page;
  template.scriptUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle(`HSE Portal — ${page}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename, context) {
  let cleanPath = filename.startsWith('Views/') ? filename : 'Views/' + filename;
  let template;
  try {
    template = HtmlService.createTemplateFromFile(cleanPath);
  } catch (e) {
    template = HtmlService.createTemplateFromFile(filename);
  }

  if (context && typeof context === 'object') {
    Object.keys(context).forEach(key => {
      template[key] = context[key];
    });
  }
  return template.evaluate().getContent();
}
