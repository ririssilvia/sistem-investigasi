function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  let page = (e && e.parameter && e.parameter.page) || 'Dashboard'; 
  const id = e && e.parameter && e.parameter.id;

  // Hapus prefix folder jika terikut dari parameter
  page = page.replace(/^Views\//, '');

  const session = token ? getValidSession(token) : null;

  if (!session) {
    const template = HtmlService.createTemplateFromFile('Views/Login');
    template.scriptUrl = ScriptApp.getService().getUrl(); 
    return template.evaluate()
      .setTitle('Investigasi Pertambangan Batubara — Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

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
  template.token = token;
  template.idInsiden = id || '';
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