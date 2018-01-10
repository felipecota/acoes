function monthDiff(d1, d2) { 
    var months = -1;   
    var dInicio = ( d1 == 0 ? new Date() : new Date(d1));
    var dFim = ( d2 == 0 ? new Date() : new Date(d2));
    while (dInicio < dFim) {
      months += 1;
      dInicio.setMonth(dInicio.getMonth() + 1);
    }
    if (months == 0)
      months = 1;
    return months; 
}

function automático() {
  var d = new Date();
  var h = d.getHours();
  var m = d.getMinutes();
  var w = d.getDay();

  if (w > 0 && w < 6 && h > 8 && h < 20) { 
     var ss = SpreadsheetApp.getActiveSpreadsheet();   
     var sheet = ss.getSheets()[0];           
     var dataOperacoes = [];
     for (var c = 12 ; c<=18 ; c++) {
       dataOperacoes.push(sheet.getRange("B"+c).getValue());
     }    
     dataOperacoes.sort();
     for (var c = 12 ; c<=18 ; c++) {
       var campo = sheet.getRange("B"+c).getValue();
       if (Utilities.formatDate(campo, "GMT", "yyyyMMdd HH:mm:ss") == Utilities.formatDate(dataOperacoes[0], "GMT", "yyyyMMdd HH:mm:ss"))
       {
         var t = buscaCotacaoMSN(sheet.getRange("A"+c).getValue()); 
         var d = new Date(t["Ld"].substring(6,10),Number(t["Ld"].substring(3,5))-1,t["Ld"].substring(0,2),t["Lt"].substring(0,2),t["Lt"].substring(3,5),t["Lt"].substring(6,8));       
         sheet.getRange("B"+c).setValue(d); 
         sheet.getRange("D"+c).setValue(t["Lp"].toString().replace('.',','));                                 
       }        
     }
  }
}

function autualizatudo() { 
  var d = new Date();
  var h = d.getHours();
  var m = d.getMinutes();
  var w = d.getDay();

  if (w > 0 && w < 6 && h > 8 && h < 20) {
     var ss = SpreadsheetApp.getActiveSpreadsheet();   
     var sheet = ss.getSheets()[0];           
     for (var c = 12 ; c<=18 ; c++) {
       var t = buscaCotacaoMSN(sheet.getRange("A"+c).getValue()); 
       var d = new Date(t["Ld"].substring(6,10),Number(t["Ld"].substring(3,5))-1,t["Ld"].substring(0,2),t["Lt"].substring(0,2),t["Lt"].substring(3,5),t["Lt"].substring(6,8));       
       sheet.getRange("B"+c).setValue(d);
       sheet.getRange("D"+c).setValue(t["Lp"].toString().replace('.',','));                        
     }
  }
}

function buscaCotacaoMSN(ativo) {
  var caminho = "https://finance.services.appex.bing.com/Market.svc/ChartAndQuotes?symbols=56.1."+ativo+".BSP&chartType=1d&isETF=false&iseod=False&lang=pt-BR&isCS=false&isVol=true";
  var json = JSON.parse(UrlFetchApp.fetch(caminho).getContentText());
  return {
    Lp: json[0]["Quotes"]["Lp"],
    Ld: json[0]["Quotes"]["Ld"],
    Lt: json[0]["Quotes"]["Lt"]
  };
}

function calculaPrecoMedio(ativo, data_limite, modo) {
  //data_limite = '06/01/2017';
  //ativo = "TOTS3";
  //modo = 2;
    
  if (data_limite == undefined)
    data_limite = new Date();
  
  data_limite = Utilities.formatDate(new Date(data_limite), "GMT", "yyyyMMdd");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();   
   
  // Variáveis de controle
  var dadosOperacoes = {};
  var dataOperacoes = [];
  var c = 0;  
  
  // Adiciona operações ativas
  var dataSheet = ss.getSheetByName("Ativos");    
  var data = dataSheet.getRange("A2:G107").getValues();
  for (var i = 0; i < data.length; ++i) {
    var rowData = data[i];
    if (rowData[0] == ativo) {
      var dt = Utilities.formatDate(new Date(rowData[1]), "GMT", "yyyyMMdd");
      if (dt <= data_limite) {
        if (dadosOperacoes[dt]) {
          dadosOperacoes[dt][0] = dadosOperacoes[dt][0]+rowData[2];
          if (modo == 2)
            dadosOperacoes[dt][1] = dadosOperacoes[dt][1]+rowData[3]+rowData[4]-rowData[6];
          else
            dadosOperacoes[dt][1] = dadosOperacoes[dt][1]+rowData[3]+rowData[4];            
        } else {
          dataOperacoes.push(dt);
          dadosOperacoes[dt] = {};      
          dadosOperacoes[dt][0] = rowData[2];
          if (modo == 2)
            dadosOperacoes[dt][1] = rowData[3]+rowData[4]-rowData[6];
          else
            dadosOperacoes[dt][1] = rowData[3]+rowData[4];            
          dadosOperacoes[dt][2] = "C";
        }
        c++;
      }
    }
  } 
  
  // Acrescento histórico de compra
  dataSheet = ss.getSheetByName("Historico");    
  data = dataSheet.getRange("A2:M107").getValues();  
  for (var i = 0; i < data.length; ++i) {
    var rowData = data[i];
    if (rowData[0] == ativo) {
      var dt = Utilities.formatDate(new Date(rowData[1]), "GMT", "yyyyMMdd");
      if (dt < data_limite) {
        if (dadosOperacoes[dt]) {
          dadosOperacoes[dt][0] = dadosOperacoes[dt][0]+rowData[3];
          if (modo == 2)
            dadosOperacoes[dt][1] = dadosOperacoes[dt][1]+rowData[4]+rowData[6]-rowData[12];                    
          else 
            dadosOperacoes[dt][1] = dadosOperacoes[dt][1]+rowData[5]+rowData[6];                                
        } else {      
          dataOperacoes.push(dt);
          dadosOperacoes[dt] = {};
          dadosOperacoes[dt][0] = rowData[3];                         
          if (modo == 2)
              dadosOperacoes[dt][1] = rowData[4]+rowData[6]-rowData[12];                        
            else
              dadosOperacoes[dt][1] = rowData[5]+rowData[6];            
          dadosOperacoes[dt][2] = "C";           
        }
        c++;
      }
    }
  }  
  
  // Acrescento histórico de venda
  for (var i = 0; i < data.length; ++i) {
    var rowData = data[i];
    if (rowData[0] == ativo) {
      var dt = Utilities.formatDate(new Date(rowData[2]), "GMT", "yyyyMMdd");
      if (dt < data_limite) {      
        if (dadosOperacoes[dt]) {        
          dadosOperacoes[dt][0] = dadosOperacoes[dt][0]+rowData[3];
        } else {      
          dataOperacoes.push(dt);         
          dadosOperacoes[dt] = {};      
          dadosOperacoes[dt][0] = rowData[3];
          dadosOperacoes[dt][1] = 0;
          dadosOperacoes[dt][2] = "V";      
        }
        c++;
      }
    }
  }    
  
  // Ordeno pela data da compra
  dataOperacoes.sort();  
  
  // Calculo Preço Médio
  var pm = 0;
  var saldo = 0;
  for (var i = 0; i < dataOperacoes.length; ++i) {
    var valor = pm*saldo;
    if (dadosOperacoes[dataOperacoes[i]][2] == "C") {
      pm = (valor+dadosOperacoes[dataOperacoes[i]][1])/(dadosOperacoes[dataOperacoes[i]][0]+saldo);
      pm = Math.round(pm*10000000)/10000000;
      saldo = dadosOperacoes[dataOperacoes[i]][0]+saldo;
    } else {
      saldo = saldo - dadosOperacoes[dataOperacoes[i]][0];
    }
    if(saldo == 0)
      pm = 0;
  }  
  
  return pm;
}
