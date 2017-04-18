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

/*
function retornaValor(ativo) {
  var parsedBovespaResponse = buscaCotacaoBovespa(ativo);
  if (parsedBovespaResponse != undefined && parsedBovespaResponse.getRootElement().getChildren()[0] != undefined) {
    var ultimaCotacao = parsedBovespaResponse.getRootElement().getChildren()[0].getAttribute("Ultimo").getValue();  
    var dataCotacao = parsedBovespaResponse.getRootElement().getChildren()[0].getAttribute("Data").getValue();
    var oscilacaoCotacao = parsedBovespaResponse.getRootElement().getChildren()[0].getAttribute("Oscilacao").getValue();
    return ultimaCotacao+';'+dataCotacao+';'+oscilacaoCotacao;
  } else {
    return undefined;
  }
}
*/

function automático() {
  var d = new Date();
  var h = d.getHours();
  var m = d.getMinutes();
  var w = d.getDay();

  if (w > 0 && w < 6 && h > 9 && h < 20) { 
     var ss = SpreadsheetApp.getActiveSpreadsheet();   
     var sheet = ss.getSheets()[0];           
     var dataOperacoes = [];
     for (var c = 12 ; c<=18 ; c++) {
       var valor_texto = sheet.getRange("B"+c).getValue(); 
       var valor = new Date(valor_texto.split(" ")[0].split("/")[2],valor_texto.split(" ")[0].split("/")[1]-1,valor_texto.split(" ")[0].split("/")[0]);
       valor.setHours(valor_texto.split(" ")[1].split(":")[0],valor_texto.split(" ")[1].split(":")[1],valor_texto.split(" ")[1].split(":")[2],00);
       valor = Utilities.formatDate(valor, "GMT", "yyyyMMdd HH:mm:ss");
       dataOperacoes.push(valor);
     }    
     dataOperacoes.sort();
     for (var c = 12 ; c<=18 ; c++) {
       var valor_texto = sheet.getRange("B"+c).getValue(); 
       var valor = new Date(valor_texto.split(" ")[0].split("/")[2],valor_texto.split(" ")[0].split("/")[1]-1,valor_texto.split(" ")[0].split("/")[0]);
       valor.setHours(valor_texto.split(" ")[1].split(":")[0],valor_texto.split(" ")[1].split(":")[1],valor_texto.split(" ")[1].split(":")[2],00);
       var campo = Utilities.formatDate(valor, "GMT", "yyyyMMdd HH:mm:ss");       
       if (campo == dataOperacoes[0])
       {
         var d = new Date();
         var timeHoje = Utilities.formatDate(d, "GMT-03:00", "dd/MM/yyyy HH:mm:ss");         
         sheet.getRange("B"+c).setValue(timeHoje);         
         var valorCotacao = buscaCotacaoYahoo(sheet.getRange("A"+c).getValue());
         sheet.getRange("D"+c).setValue(valorCotacao.replace('.',','));          
         /*
         var retorno = retornaValor(sheet.getRange("A"+c).getValue());
         if (retorno != undefined) {          
           var valores = retorno.split(";");
           sheet.getRange("B"+c).setValue(valores[1]);         
           sheet.getRange("D"+c).setValue(valores[0]);         
         } 
         */
       }        
     }
  }
}

function autualizatudo() { 
  var d = new Date();
  var h = d.getHours();
  var m = d.getMinutes();
  var w = d.getDay();

  if (w > 0 && w < 6 && h > 9 && h < 20) {
     var ss = SpreadsheetApp.getActiveSpreadsheet();   
     var sheet = ss.getSheets()[0];           
     for (var c = 12 ; c<=18 ; c++) {
       var d = new Date();
       var timeHoje = Utilities.formatDate(d, "GMT-03:00", "dd/MM/yyyy HH:mm:ss");         
       sheet.getRange("B"+c).setValue(timeHoje);         
       var valorCotacao = buscaCotacaoYahoo(sheet.getRange("A"+c).getValue());
       sheet.getRange("D"+c).setValue(valorCotacao.replace('.',','));                 
       /*
       var retorno = retornaValor(sheet.getRange("A"+c).getValue());
       if (retorno != undefined) {          
         var valores = retorno.split(";");
         sheet.getRange("B"+c).setValue(valores[1]);         
         sheet.getRange("D"+c).setValue(valores[0]);         
         //sheet.getRange("E"+c).setValue(valores[2]+"%");         
       }
       */
     }
  }
}

function buscaCotacaoYahoo(ativo) {
  var caminho = "http://finance.yahoo.com/d/quotes.csv?s="+ativo+".SA&f=l1";
  var csv = Utilities.parseCsv(UrlFetchApp.fetch(caminho).getContentText());
  return csv[0][0];
}

/*
function buscaCotacaoBovespa(ativo) {
  var caminho = "www.bmfbovespa.com.br/Pregao-Online/ExecutaAcaoAjax.asp?CodigoPapel="+ativo;    
  try {
    return XmlService.parse(UrlFetchApp.fetch(caminho).getContentText());
  } catch (e) {
    return undefined;
  }
}
*/

function calculaPrecoMedio(ativo, data_limite, modo) {
  //data_limite = '02/06/2016';
  //ativo = "ARZZ3";
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
