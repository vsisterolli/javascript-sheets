var sheet = SpreadsheetApp.getActiveSpreadsheet();

function verificar_correcoes(atividade, intervalo, start_line) {
    
    let aba = sheet.getSheetByName(atividade);
    SpreadsheetApp.setActiveSheet(aba);

    atividades_realizadas = aba.getRange(intervalo).getValues();

    let = todos_corrigidos = true;
    for(let i = 0; i < atividades_realizadas.length; i++)
      if(atividades_realizadas[i][0] != "" && atividades_realizadas[i][3] == "") {
        console.log("Um " + atividade + " posto na linha " + (i+start_line) + " ainda não foi corrigido.");
        todos_corrigidos = false;
      }

    return todos_corrigidos;

}

function foram_corrigidos() {
  let corrigidos = true;
  
  atividades = ["Testes Práticos ®", "Reforços ®", "Auxílios externos ®", "Erros ortográficos ®", "Fiscalizações ®", "Destaques ®"];
  intervalos = ["H9:K", "E12:M", "E10:H", "G10:J", "H10:P", "I12:L"];
  start_line = [9, 12, 10, 10, 10, 12];

  for(let i = 0; i < atividades.length; i++)
    corrigidos = verificar_correcoes(atividades[i], intervalos[i], start_line[i])
  
  return corrigidos;
}

function atualizada_hoje() {
  
  let aba = sheet.getSheetByName('Contabilização ®');
  SpreadsheetApp.setActiveSheet(aba);

  let today = aba.getRange('B5:B5').getValue(); // B5 guarda o resultado da fórmula =TODAY();
  let ultima_contabilização = planilha.getRange('CE1:CE1'); // CE1 guarda a data da última contabilização realizada;

  if(today == ultima_contabilização) {
      console.log("A planilha já foi contabilizada hoje.");
      return true;
  }
  else {
      console.log("Ela ainda não foi contabilizada hoje! Iniciando contabilização...")
      ultima_contabilização.setValue(today);
      return false;
  }
  
}

function atualizar_semana(start_week, range2update, id_semana) {
    
    let aba = sheet.getSheetByName('Contabilização ®');
    SpreadsheetApp.setActiveSheet(aba);

    start_week = aba.getRange(start_week).getValue();

    let end_week = start_week;
    end_week.setHours(168); // 7 dias após start_week;
    let today = aba.getRange('B5:B5').getValue(); // B5 guarda o resultado da fórmula =TODAY();

    if(today >= end_week) { // se a semana conferida já acabou
        
        console.log("A semana " + id_semana + " acabou e os valores já estão sendo substituidos!");
        
        let content = aba.getRange(range2update);
        let values = content.getValues();

        console.log(values);
        content.setValues(values);

    }
    else 
        console.log("A semana " + id_semana + " ainda não acabou!");
    

    return;

}

function iterateSemanas() {
   let start = ["D6:M6", "O6:X6", "Z6:AI6", "AV6:BE6"]
   let range = ["D9:M", "O9:X", "Z9:AI", "AV9:BE"];

   for(let i = 0; i < 4; i++)
    atualizar_semana(start[i], range[i], i+1);
}


function contabilizar() {


  // verificar correção
  
  if(!foram_corrigidos()) {
    console.log("A planilha terminará de contabilizar assim que todas as atividades estiverem corrigidas");
    return;
  }
  else
    console.log("Todos os cursos foram corrigidos!");

  // verificar se já foi atualizada hoje
  if(atualizada_hoje()) {
    console.log("A planilha já foi contabilizada");
    return;
  }

  // iniciar atualização
  iterateSemanas();

}

