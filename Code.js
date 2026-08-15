function doGet(e) {
  const token = e && e.parameter && e.parameter.token;
  const page = (e && e.parameter && e.parameter.page) || 'Dashboard'; 
  const id = e && e.parameter && e.parameter.id;

  const session = token ? getValidSession(token) : null;

  if (!session) {
    const template = HtmlService.createTemplateFromFile('Login');
    template.scriptUrl = ScriptApp.getService().getUrl(); 
    return template.evaluate()
      .setTitle('Investigasi Pertambangan Batubara — Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const template = HtmlService.createTemplateFromFile(page);
  template.userName = session.Nama || session.Username;
  template.userRole = session.Role;
  template.userSite = session.Site;
  template.token = token;
  template.idInsiden = id || '';
  template.page = page; // <-- Tambahkan baris ini agar sidebar tahu halaman aktif
  template.scriptUrl = ScriptApp.getService().getUrl(); 

  return template.evaluate()
    .setTitle(`HSE Portal — ${page}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename, context) {
  const template = HtmlService.createTemplateFromFile(filename);
  
  // Jika context dikirimkan, salin variabelnya
  if (context) {
    Object.keys(context).forEach(key => {
      template[key] = context[key];
    });
  }
  
  return template.evaluate().getContent();
}