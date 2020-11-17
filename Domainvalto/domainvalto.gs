function Domainvalto() {

  var mit = "sulinet.hu";
  var mire = "edu.hu";
  
  // Felhasználók számának meghatározása
  var lista;
  var par = { customer: "my_customer",
              maxResults: 500,
              pageToken: "" };
  var userMAX = 0;
  SpreadsheetApp.getActiveSheet().getRange(1,1).setValue("Felhasználók számának meghatározása...");
  SpreadsheetApp.flush();
  do {
    lista = AdminDirectory.Users.list(par);
    par.pageToken = lista.nextPageToken;
    userMAX+=lista.users.length;
  } while (par.pageToken);
  
  // Progressbar megoldása Sparkline() függvénnyel
  SpreadsheetApp.getActiveSheet().setColumnWidth(1, 300);
  SpreadsheetApp.getActiveSheet().setColumnWidth(2, 300);
  SpreadsheetApp.getActiveSheet().getRange(1,1).setValue("Felhasználók e-mail címének módosítása");
  SpreadsheetApp.getActiveSheet().getRange(1,2).setFormula('=SPARKLINE(C1;{"charttype"\\"bar";"max"\\'+userMAX+';"color1"\\"#960000"})');
  SpreadsheetApp.flush();

  //--- Felhasználók e-mail címeinek módosítása

  var sor=1;
  var db=0;
  var oldEmail;
  par.pageToken = "";
  par.maxResults=50;

  do {

    lista = AdminDirectory.Users.list(par);

    par.pageToken = lista.nextPageToken;
    db+=lista.users.length;

    for (var i = 0; i < lista.users.length; i++){
      if(lista.users[i].primaryEmail.indexOf("teszt")>=0){
        oldEmail=lista.users[i].primaryEmail;
        lista.users[i].aliases.push(oldEmail);
        var felh = { primaryEmail: lista.users[i].primaryEmail.replace(mit, mire),
                     aliases: lista.users[i].aliases };
        AdminDirectory.Users.update(felh, lista.users[i].id);
        
        // Eredmény ellenőrző rész
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,1).setValue(lista.users[i].primaryEmail);
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,2).setValue(lista.users[i].id);
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,3).setValue(lista.users[i].aliases);
        sor++;
      }
    }
    SpreadsheetApp.getActiveSheet().getRange(1,3).setValue(db);
    SpreadsheetApp.flush();

  } while (par.pageToken);
  


  //--- Google csoportok e-mail címeinek módosítása

  sor=1;
  par.pageToken = "";
  par.maxResults=100;

  do {
    
    lista = AdminDirectory.Groups.list(par);

    par.pageToken = lista.nextPageToken;

    for (var i = 0; i < lista.groups.length; i++){
      if(lista.groups[i].name.indexOf("teszt")>=0){
        oldEmail=lista.groups[i].email;
        var csop = { email: lista.groups[i].email.replace(mit, mire),
                     alias: oldEmail };
        AdminDirectory.Groups.update(csop, lista.groups[i].id);
        
        // Eredmény ellenőrző rész
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,1).setValue(lista.groups[i].name);
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,2).setValue(lista.groups[i].id);
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,3).setValue(lista.groups[i].email);
        //SpreadsheetApp.getActiveSheet().getRange(sor+2,4).setValue(lista.groups[i].aliases);
        sor++;

      }
    }

  } while (par.pageToken);  

}
