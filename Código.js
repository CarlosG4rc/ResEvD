function onFormSubmit(e) {
  var nombre = e.values[1];
  var correo = e.values[2];
  var padre = 'jbarrientos@ifp.mx'
  var direc = 'gjuarez@ifp.mx'
//----------------------------------------------------------Creacion de ssnew mostrará dashboard
  var ssnew = SpreadsheetApp.create(nombre + ' Resultados Evaluacion Docente').addViewer(correo).addViewer(direc).addEditor(padre);
  var sheet = ssnew.getSheetByName('Hoja 1').setHiddenGridlines(true);
//----------------------------------------------------------Id de SS usadas
  var ssresp1 = SpreadsheetApp.openById("1xtD2DMPqaE-3vBr1ua4GCh87iq3bWhD34oelqCJzHlI"); //SS evaluación 1o
  var ssresp3 = SpreadsheetApp.openById("1tWTbqt5_rkRd1QAw7hoLOAl4LsavihpABaZOgdo183w"); //SS evaluación 3o
  var bd = SpreadsheetApp.openById("1wwlLUNX7t58C6vceM0lYFravJD6vqM0NjMtJzP9iJTU"); //SS Base de datos profesores
//----------------------------------------------------------Validar nombre de profesor y correo
  var lastrowbd = bd.getDataRange().getNumRows();
  var column = bd.getDataRange();
  var value = column.getValues();
  for(var i = 0; i < lastrowbd; i++)//Recorremos todas las filas 
  {
    if(value[i][8] == nombre)//Comparamos nombre introducido con base de datos
    {
       var email = value[i] && value[i][2];//guardamos el correo
       ssnew.getRange("A1").setValue(email);
       i = lastrowbd + 1;//Salimos del loop
    }
  }
//----------------------------------------------------------Filtro información por profesor en ssresp1 y ssresp3 tenemos una hoja de cada profesor en su SS correspondiente
//----------------------------------------------------------Damos forma a tabulación de datos numéricos
  var tab = ssresp3.getSheetByName('Respuestas3').getRange("F1:U1").getValues();
  var better = ssresp3.getSheetByName('Respuestas3').getRange("V1").getValues();
  ssnew.getRange("A1:Z1").setBackgroundRGB(0, 0, 128).setFontColor("white");
  ssnew.getRange("A3:Z3").setBackgroundRGB(0, 0, 128);
  ssnew.getRange("A8:Z8").setBackgroundRGB(0, 0, 128);
  ssnew.getRange("B3:B12").setBackgroundRGB(0, 0, 128).setFontColor("white");
  ssnew.getRange("A14:Z44").setBackgroundRGB(0, 0, 128);
  ssnew.getRange("C1").setValue(nombre).setFontSize("20");//Nombre del profesor
  ssnew.getRange("C3:R3").setValues(tab).setFontColor("white");//Tabla 1
  ssnew.getRange("A3").setValue("Primera evaluación").setFontColor("white");
  ssnew.getRange("C8:R8").setValues(tab).setFontColor("white");//Tabla2
  ssnew.getRange("T8").setValues(better).setFontColor("white");//Pregunta si hubo mejora
  ssnew.getRange("A8").setValue("Segunda evaluación").setFontColor("white");
  var datos1 = ssresp1.getSheetByName(nombre).getRange("E1:U4").getValues();//Guardamos datos de la primera evaluación del profesor señalado
  ssnew.getRange("B4:R7").setValues(datos1);//Mostramos datos en la nueva SS
  var datos2 = ssresp3.getSheetByName(nombre).getRange("E1:U4").getValues();//Guardamos datos de la segunda evaluación del profesor señalado 
  ssnew.getRange("B9:R12").setValues(datos2);//Mostramos datos en la nueva SS
  var mejora = ssresp3.getSheetByName(nombre).getRange("V1:V2").getValues();//Copiamos los valores de grafica de mejora
  ssnew.getRange("U9:U10").setValues(mejora);//Pegamos valores para grafica de mejora
  ssnew.getRange("T9").setValue("Sí");//Membrete para grafica de mejora SI
  ssnew.getRange("T10").setValue("No");//Membrete para grafica de mejora NO
  ssnew.getRange("T9:U9").setFontColor("white").setBackground("green");
  ssnew.getRange("T10:U10").setFontColor("white").setBackground("red");
  var prom1 = ssresp1.getSheetByName(nombre).getRange("W1:W2").getValues();//Copiar promedio desde ssresp1 y ssresp3
  ssnew.getRange("S3:S4").setValues(prom1);//Pegamos datos de promedio
  var prom2 = ssresp3.getSheetByName(nombre).getRange("X1:X2").getValues();
  ssnew.getRange("S8:S9").setValues(prom2);
  ssnew.getRange("S3").setFontColor("white");
  ssnew.getRange("S8").setFontColor("white");
  //---------------------------------------------Mostramos graficos
   var hoja = ssnew.getSheetByName('Hoja 1');
   
   var grafico1 = hoja.newChart()//Construimos grafica 1
       .setChartType(Charts.ChartType.COLUMN)
       .addRange(hoja.getRange("B3:B7"))
       .addRange(hoja.getRange("C3:R7"))
       .setPosition(14,1,0,0)
       .setOption('width', 800)
       .setOption('height', 640)
       .setNumHeaders(1)
       .setOption("legend", {position: "top"})
       .setOption('title','Primera Evaluación Docente')
       .setTransposeRowsAndColumns(true)
       .build();
       
   var grafico2 = hoja.newChart()//Construimos grafica 2
       .setChartType(Charts.ChartType.COLUMN)
       .addRange(hoja.getRange("B8:B12"))
       .addRange(hoja.getRange("C8:R12"))
       .setPosition(14,9,0,0)
       .setOption('width', 800)
       .setOption('height', 640)
       .setNumHeaders(1)
       .setOption("legend", {position: "top"})
       .setOption('title','Segunda Evaluación Docente')
       .setTransposeRowsAndColumns(true)
       .build();

   var grafico3 = hoja.newChart()
       .setChartType(Charts.ChartType.PIE)
       .addRange(hoja.getRange("T9:T10"))
       .addRange(hoja.getRange("U9:U10"))
       .setPosition(46,6,0,0)
       .setOption("legend", {position: "right"})
       .setOption('title', '¿Ha mejorado la clase? Respecto a la evaluación anterior')
       .build();
  
//Graficar promedio
  var grafico4 = hoja.newChart()
       .setChartType(Charts.ChartType.GAUGE)
       .addRange(hoja.getRange("S3:S4"))
       .setPosition(65,6,0,0)
       .setOption('height', 300)
       .setOption('width', 300)
       .setOption('title', 'Promedio Anterior')
       .setOption('max',10)
       .build();
       
    var grafico5 = hoja.newChart()
       .setChartType(Charts.ChartType.GAUGE)
       .addRange(hoja.getRange("S8:S9"))
       .setPosition(65,9,0,0)
       .setOption('height', 300)
       .setOption('width', 300)
       .setOption('title', 'Promedio Actual')
       .setOption('max',10)
       .build();     
  
   hoja.insertChart(grafico1);
   hoja.insertChart(grafico2);
   hoja.insertChart(grafico3);
   hoja.insertChart(grafico4);
   hoja.insertChart(grafico5);
	  
//----------------------------------------------------------Indentifico ultimo renglon de información 
  var lastrow3 = ssresp3.getSheetByName(nombre).getDataRange().getNumRows();
  var lastrownew = lastrow3 + 76;  
//----------------------------------------------------------Copio y pego recomendación escrita por alumnos 
  var opcopy = ssresp3.getSheetByName(nombre).getRange("W5:W" + lastrow3).getValues();
  var opcopy2 = ssresp3.getSheetByName(nombre).getRange("Y5:Y" + lastrow3).getValues();
  ssnew.getRange("A80:Z80").setBackgroundRGB(0, 0, 128);
  ssnew.getRange("E80").setValue("Propuestas para mejorar").setFontColor("white");
  ssnew.getRange("E81:E" + lastrownew).setValues(opcopy);
  ssnew.getRange("H80").setValue("Comentarios Extra").setFontColor("white");
  ssnew.getRange("H81:H" + lastrownew).setValues(opcopy2);
  var url = ssnew.getUrl();
  var body = 'Evaluación docente ' + url;
  MailApp.sendEmail(correo,'Resultado de Evaluación Docente' + nombre, body);
        
}
