#######################################################################################
# Google Workspace/G Suit kurzuslistázó
#
# Ez a script kilistázza az összes kurzust és az abban szereplő tanárokat és diákokat.
#
# Készítette: Venczel József
#
# Debrecen, 2020. 11. 18.
#######################################################################################



function Kurzuslistazo() {

  var sor=1;
  
  do {

    var par = { pageSize: 10,
                pageToken: "" };

    var lista = Classroom.Courses.list(optionalArgs);
   
    par.pageToken = lista.nextPageToken;

    for (var i = 0; i < lista.courses.length; i++){
      if(lista.courses[i].name.indexOf("teszt")<0){
        SpreadsheetApp.getActiveSheet().getRange(sor,1).setValue(lista.courses[i].name);
        sor++;

        tanarok=Classroom.Courses.Teachers.list(lista.courses[i].id);
        if(tanarok.teachers)
          for(var j = 0; j<tanarok.teachers.length; j++){
            SpreadsheetApp.getActiveSheet().getRange(sor,2).setValue(tanarok.teachers[j].profile.name.fullName);
            sor++;
          }
      
        var nextPageToken2 = "";
        do{
          var studentsArgs = { pageSize: 30,
                               pageToken: nextPageToken2};  
          diakok=Classroom.Courses.Students.list(lista.courses[i].id, studentsArgs);
          if(diakok.students)
            for(var j = 0; j<diakok.students.length; j++){
              SpreadsheetApp.getActiveSheet().getRange(sor,2).setValue(diakok.students[j].profile.name.fullName);
              sor++;
            }

          nextPageToken2 = diakok.nextPageToken;
        } while (nextPageToken2)
      }
    }

  } while (par.pageToken);
}