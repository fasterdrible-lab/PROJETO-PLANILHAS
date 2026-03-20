// ============================================================
// ESTADO DA CONVERSA
// ============================================================
function _getEstado(chatId) {
  var raw = PropertiesService.getScriptProperties()
              .getProperty("estado_"+chatId);
  return raw ? JSON.parse(raw) : { etapa: "MENU" };
}
function _setEstado(chatId, estado) {
  PropertiesService.getScriptProperties()
    .setProperty("estado_"+chatId, JSON.stringify(estado));
}
function _limparEstado(chatId) {
  PropertiesService.getScriptProperties()
    .deleteProperty("estado_"+chatId);
}

// ============================================================
// WEBHOOK — apenas UMA definição
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var update   = JSON.parse(e.postData.contents);
    var updateId = String(update.update_id || "");
    var prop     = PropertiesService.getScriptProperties();

    // Anti-duplicata
    if (updateId && prop.getProperty("upd_" + updateId)) {
      console.log("⚠️ Update " + updateId + " já processado — ignorado");
      return ContentService.createTextOutput("OK");
    }

    // Lock máx 8 segundos
    if (!lock.tryLock(8000)) {
      console.log("⚠️ Lock ocupado — ignorando " + updateId);
      return ContentService.createTextOutput("OK");
    }

    // Dupla verificação após lock
    if (updateId && prop.getProperty("upd_" + updateId)) {
      console.log("⚠️ Update " + updateId + " já processado (pós-lock)");
      return ContentService.createTextOutput("OK");
    }

    if (updateId) prop.setProperty("upd_" + updateId, "1");
    _processarUpdate(update);

  } catch(err) {
    console.log("❌ doPost: " + err.message);
  } finally {
    try { lock.releaseLock(); } catch(e2) {}
  }
  return ContentService.createTextOutput("OK");
}

function doGet(e) {
  return ContentService.createTextOutput("✅ Bot FROTA 17ºGB online!");
}

// ============================================================
// PROCESSAMENTO
// ============================================================
function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName;

  if (callback) {
    chatId   = callback.message.chat.id.toString();
    texto    = callback.data;
    fromName = callback.from.first_name || "Auxiliar";
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId   = update.message.chat.id.toString();
    texto    = update.message.text || "";
    fromName = update.message.from.first_name || "Auxiliar";
  } else {
    return;
  }

  // Verifica chat autorizado
  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName);
}

function _responderCallback(callbackId) {
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/answerCallbackQuery",
    { method:"post", contentType:"application/json",
      payload: JSON.stringify({callback_query_id: callbackId}),
      muteHttpExceptions: true }
  );
}

// ============================================================
// ROTEADOR
// ============================================================
function _rotear(chatId, texto, fromName) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName);
    return;
  }

  // Qualquer etapa pode voltar ao menu
  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName);
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName);
  }
}

// ============================================================
// MENU PRINCIPAL
// ============================================================
function _enviarMenu(chatId, nome) {
  _setEstado(chatId, { etapa:"MENU" });
  var msg = "👋 Olá, *"+(nome||"Auxiliar")+"*!\n\n"
    + "📋 *FROTA 17ºGB — Menu Principal*\n\n"
    + "O que deseja fazer?";
  var teclado = { inline_keyboard: [
    [{ text:"🚒 Atualizar Viatura",  callback_data:"MENU_VTR"    }],
    [{ text:"📋 Gerenciar Tarefas",  callback_data:"MENU_TAR"    }],
    [{ text:"📊 Ver Resumo Frota",   callback_data:"MENU_RESUMO" }],
    [{ text:"⛽ Ver Alertas Abast.", callback_data:"MENU_ABAST"  }]
  ]};
  _enviarTeclado(chatId, msg, teclado);
}

function _tratarMenu(chatId, texto, fromName) {
  switch(texto) {
    case "MENU_VTR":    _iniciarVtr(chatId);     break;
    case "MENU_TAR":    _iniciarTarefas(chatId); break;
    case "MENU_RESUMO": _enviarResumo(chatId);   break;
    case "MENU_ABAST":  _enviarAbast(chatId);    break;
    default: _enviarMenu(chatId, fromName);
  }
}

// ============================================================
// FLUXO — ATUALIZAR VIATURA
// ============================================================
function _iniciarVtr(chatId) {
  _setEstado(chatId, { etapa:"VTR_ESCOLHER_ABA" });
  var teclado = { inline_keyboard: [
    [{ text:"1️⃣ 1SGB", callback_data:"ABA_1SGB" },
     { text:"2️⃣ 2SGB", callback_data:"ABA_2SGB" }],
    [{ text:"🔙 Voltar", callback_data:"VOLTAR_MENU" }]
  ]};
  _enviarTeclado(chatId, "🚒 Qual subgrupamento?", teclado);
}

function _tratarEscolherAba(chatId, texto, estado) {
  if (texto !== "ABA_1SGB" && texto !== "ABA_2SGB") {
    _iniciarVtr(chatId); return;
  }

  var nomeAba = texto === "ABA_1SGB" ? ABA_1SGB : ABA_2SGB;
  var vtrs    = _listarVtrsAbaCached(nomeAba);

  if (!vtrs || !vtrs.length) {
    _enviar(chatId, "⚠️ Nenhuma viatura encontrada em: " + nomeAba);
    return;
  }

  _setEstado(chatId, { etapa:"VTR_ESCOLHER_VTR", aba:nomeAba, abaKey:texto, vtrs:vtrs });

  var linhas = [];
  for (var i = 0; i < vtrs.length; i += 2) {
    var linha = [{ text:vtrs[i].prefixo, callback_data:"VTR_"+i }];
    if (vtrs[i+1]) linha.push({ text:vtrs[i+1].prefixo, callback_data:"VTR_"+(i+1) });
    linhas.push(linha);
  }
  linhas.push([{ text:"🔙 Voltar", callback_data:"VOLTAR_MENU" }]);

  _enviarTeclado(chatId,
    "🚒 *"+nomeAba+"* — "+vtrs.length+" viaturas\nSelecione:",
    { inline_keyboard: linhas });
}

function _tratarEscolherVtr(chatId, texto, estado) {
  if (texto === "VOLTAR_ABA") {
    _iniciarVtr(chatId); return;
  }

  if (texto.indexOf("VTR_") !== 0) {
    _iniciarVtr(chatId); return;
  }

  var idx = parseInt(texto.replace("VTR_",""), 10);
  if (isNaN(idx) || !estado.vtrs || !estado.vtrs[idx]) {
    _enviar(chatId, "⚠️ Viatura inválida. Selecione novamente.");
    Utilities.sleep(300);
    // Reenvia lista de viaturas
    _tratarEscolherAba(chatId, estado.abaKey, { etapa:"VTR_ESCOLHER_ABA" });
    return;
  }

  var vtr = estado.vtrs[idx];
  _setEstado(chatId, {
    etapa:  "VTR_ESCOLHER_CAMPO",
    aba:    estado.aba,
    abaKey: estado.abaKey,
    vtrIdx: idx,
    vtr:    vtr,
    vtrs:   estado.vtrs
  });

  var info = "🚒 *"+vtr.prefixo+"* — "+vtr.placa+"\n"
    + "📍 KM atual: *"+vtr.km.toLocaleString("pt-BR")+"*\n\n"
    + "Qual campo deseja atualizar?";

  var teclado = { inline_keyboard: [
    [{ text:"💧 Próx. Troca Óleo (KM)",   callback_data:"CAMPO_OleoKm"    },
     { text:"📅 Próx. Troca Óleo (Data)", callback_data:"CAMPO_OleoData"  }],
    [{ text:"🛑 Revisão Freio (KM)",      callback_data:"CAMPO_FreioKm"   },
     { text:"🔋 Garantia Bateria (Data)", callback_data:"CAMPO_Bateria"   }],
    [{ text:"🚿 Data Lavagem",            callback_data:"CAMPO_Lavagem"   },
     { text:"🔵 Garantia VTR (Data)",     callback_data:"CAMPO_Garantia"  }],
    [{ text:"🟡 Pneu: KM Próx. Troca",   callback_data:"CAMPO_Pneu"      },
     { text:"⚙️ KM Troca Embreagem",     callback_data:"CAMPO_Embreagem" }],
    [{ text:"🔙 Voltar",                  callback_data:"VOLTAR_ABA"      }]
  ]};

  _enviarTeclado(chatId, info, teclado);
}

function _tratarEscolherCampo(chatId, texto, estado) {
  if (texto === "VOLTAR_VTR" || texto === "VOLTAR_ABA") {
    // Volta para lista de viaturas
    _setEstado(chatId, { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
    _tratarEscolherAba(chatId, estado.abaKey, { etapa:"VTR_ESCOLHER_ABA" });
    return;
  }

  var CAMPOS = {
    "CAMPO_OleoKm"   :{ nome:"Próx. Troca Óleo (KM)",   col:4,  tipo:"km",   ex:"Ex: 165000"     },
    "CAMPO_OleoData" :{ nome:"Próx. Troca Óleo (Data)", col:5,  tipo:"data", ex:"Ex: 26/06/2026" },
    "CAMPO_FreioKm"  :{ nome:"Revisão Freio (KM)",      col:6,  tipo:"km",   ex:"Ex: 170000"     },
    "CAMPO_Bateria"  :{ nome:"Data Venc. Bateria",      col:7,  tipo:"data", ex:"Ex: 26/10/2026" },
    "CAMPO_Lavagem"  :{ nome:"Data Última Lavagem",     col:12, tipo:"data", ex:"Ex: 20/03/2026" },
    "CAMPO_Garantia" :{ nome:"VTR em Garantia (Data)",  col:3,  tipo:"data", ex:"Ex: 09/03/2027" },
    "CAMPO_Pneu"     :{ nome:"Pneus: KM Próx. Troca",  col:13, tipo:"km",   ex:"Ex: 180000"     },
    "CAMPO_Embreagem":{ nome:"KM Troca Embreagem",      col:14, tipo:"km",   ex:"Ex: 200000"     }
  };

  var campo = CAMPOS[texto];
  if (!campo) {
    _enviar(chatId, "⚠️ Campo inválido.");
    return;
  }

  _setEstado(chatId, {
    etapa:    "VTR_INSERIR_VALOR",
    aba:      estado.aba,
    abaKey:   estado.abaKey,
    vtr:      estado.vtr,
    vtrIdx:   estado.vtrIdx,
    vtrs:     estado.vtrs,
    campo:    campo,
    campoKey: texto
  });

  _enviar(chatId,
    "✏️ *"+estado.vtr.prefixo+"* — "+campo.nome+"\n\n"
    +"Digite o novo valor:\n_"+campo.ex+"_\n\n"
    +"Ou /cancelar para voltar.");
}

function _tratarInserirValor(chatId, texto, estado) {
  if (texto === "/cancelar") {
    _setEstado(chatId, {
      etapa:  "VTR_ESCOLHER_CAMPO",
      aba:    estado.aba,
      abaKey: estado.abaKey,
      vtr:    estado.vtr,
      vtrIdx: estado.vtrIdx,
      vtrs:   estado.vtrs
    });
    _tratarEscolherVtr(chatId, "VTR_"+estado.vtrIdx,
      { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
    return;
  }

  var campo = estado.campo;
  var valor, valorFmt;

  if (campo.tipo === "km") {
    var n = parseInt(texto.replace(/\D/g,""), 10);
    if (isNaN(n) || n < 1000 || n > 9999999) {
      _enviar(chatId, "⚠️ KM inválido. Digite apenas números.\n_"+campo.ex+"_");
      return;
    }
    valor    = n;
    valorFmt = n.toLocaleString("pt-BR")+" km";
  } else {
    var dt = _parseDataBot(texto);
    if (!dt) {
      _enviar(chatId, "⚠️ Data inválida. Use DD/MM/AAAA.\n_Ex: 26/06/2026_");
      return;
    }
    valor    = dt;
    valorFmt = _fmtData(dt);
  }

  var ok = _gravarCampoVtr(estado.aba, estado.vtr.linha, campo.col, valor);

  if (ok) {
    _enviar(chatId,
      "✅ *Atualizado com sucesso!*\n\n"
      +"🚒 "+estado.vtr.prefixo+" — "+estado.vtr.placa+"\n"
      +"📝 "+campo.nome+": *"+valorFmt+"*");
  } else {
    _enviar(chatId, "❌ Erro ao gravar. Tente novamente.");
  }

  Utilities.sleep(300);
  _setEstado(chatId, {
    etapa:  "VTR_ESCOLHER_CAMPO",
    aba:    estado.aba,
    abaKey: estado.abaKey,
    vtr:    estado.vtr,
    vtrIdx: estado.vtrIdx,
    vtrs:   estado.vtrs
  });
  _tratarEscolherVtr(chatId, "VTR_"+estado.vtrIdx,
    { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
}

// ============================================================
// FLUXO — TAREFAS
// ============================================================
function _iniciarTarefas(chatId) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas = _lerTarefasCompleto(ss);

  var msg = "📋 *TAREFAS — "+tarefas.total+" registros*\n\n";
  if (tarefas.pendente  > 0) msg += "🔴 PENDENTE: "     +tarefas.pendente+"\n";
  if (tarefas.andamento > 0) msg += "🟡 EM ANDAMENTO: " +tarefas.andamento+"\n";
  if (tarefas.concluida > 0) msg += "✅ CONCLUÍDA: "    +tarefas.concluida+"\n";

  _setEstado(chatId, { etapa:"TAR_MENU" });

  var teclado = { inline_keyboard: [
    [{ text:"✅ Dar Baixa em Tarefa", callback_data:"TAR_BAIXA"   }],
    [{ text:"➕ Nova Tarefa",         callback_data:"TAR_NOVA"    }],
    [{ text:"📋 Listar Pendentes",    callback_data:"TAR_LISTAR"  }],
    [{ text:"🔙 Menu Principal",      callback_data:"VOLTAR_MENU" }]
  ]};

  _enviarTeclado(chatId, msg, teclado);
}

function _tratarMenuTarefas(chatId, texto, estado) {
  switch(texto) {
    case "TAR_BAIXA":  _iniciarBaixaTarefa(chatId);     break;
    case "TAR_NOVA":   _iniciarNovaTarefa(chatId);      break;
    case "TAR_LISTAR": _listarTarefasPendentes(chatId); break;
    default: _iniciarTarefas(chatId);
  }
}

function _iniciarBaixaTarefa(chatId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas   = _lerTarefasCompleto(ss);
  var pendentes = tarefas.lista.filter(function(t) {
    return t.status !== "CONCLUIDA";
  });

  if (!pendentes.length) {
    _enviar(chatId,"✅ Não há tarefas pendentes!");
    Utilities.sleep(300);
    _iniciarTarefas(chatId);
    return;
  }

  _setEstado(chatId, { etapa:"TAR_ESCOLHER_BAIXA", pendentes:pendentes });

  var linhas = [];
  var max = Math.min(pendentes.length, 20);
  for (var i = 0; i < max; i++) {
    var t   = pendentes[i];
    var lbl = (i+1)+". "+t.descricao.substring(0,30)
              +(t.descricao.length>30?"...":"")+" ["+t.status+"]";
    linhas.push([{ text:lbl, callback_data:"BAIXA_"+i }]);
  }
  linhas.push([{ text:"🔙 Voltar", callback_data:"VOLTAR_TAR" }]);

  _enviarTeclado(chatId, "✅ *Selecione a tarefa concluída:*",
    { inline_keyboard: linhas });
}

function _tratarEscolherBaixa(chatId, texto, estado) {
  if (texto === "VOLTAR_TAR") { _iniciarTarefas(chatId); return; }

  var idx = parseInt(texto.replace("BAIXA_",""), 10);
  if (isNaN(idx) || !estado.pendentes[idx]) {
    _enviar(chatId,"⚠️ Seleção inválida."); return;
  }

  var tarefa = estado.pendentes[idx];
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ok     = _darBaixaTarefa(ss, tarefa.linha);

  _enviar(chatId, ok
    ? "✅ *Tarefa concluída!*\n\n📝 "+tarefa.descricao+"\n👤 "+tarefa.responsavel
    : "❌ Erro ao dar baixa. Tente novamente.");

  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

function _iniciarNovaTarefa(chatId) {
  _setEstado(chatId, { etapa:"TAR_NOVA_PREFIXO" });
  _enviar(chatId,
    "➕ *Nova Tarefa*\n\n"
    +"🚒 Digite o *PREFIXO* da viatura:\n"
    +"_Ex: ABS-17101_\n_(ou /cancelar para voltar)_");
}

function _tratarNovaPrefixo(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }
  _setEstado(chatId, { etapa:"TAR_NOVA_DESC", prefixo:texto });
  _enviar(chatId, "📝 Digite a *DESCRIÇÃO* da tarefa:\n_(ou /cancelar)_");
}

function _tratarNovaDesc(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }
  if (texto.length < 3) { _enviar(chatId,"⚠️ Descrição muito curta."); return; }
  _setEstado(chatId, { etapa:"TAR_NOVA_RESP", prefixo:estado.prefixo, descricao:texto });
  _enviar(chatId, "👤 Quem é o *RESPONSÁVEL*?\n_(ou /cancelar)_");
}

function _tratarNovaResp(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ok = _inserirNovaTarefa(ss, estado.prefixo, estado.descricao, texto);

  _enviar(chatId, ok
    ? "✅ *Tarefa criada!*\n\n🚒 "+estado.prefixo+"\n📝 "+estado.descricao+"\n👤 "+texto+"\n🔴 Status: PENDENTE"
    : "❌ Erro ao criar tarefa.");

  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

function _listarTarefasPendentes(chatId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas   = _lerTarefasCompleto(ss);
  var pendentes = tarefas.lista.filter(function(t) {
    return t.status !== "CONCLUIDA";
  });

  if (!pendentes.length) {
    _enviar(chatId,"✅ Nenhuma tarefa pendente!");
    Utilities.sleep(300);
    _iniciarTarefas(chatId);
    return;
  }

  var msg = "📋 *TAREFAS PENDENTES ("+pendentes.length+")*\n――――――――――――\n";
  pendentes.forEach(function(t, i) {
    var icone = t.status === "EM ANDAMENTO" ? "🟡" : "🔴";
    msg += icone+" *"+(i+1)+".* "+t.descricao+"\n";
    if (t.responsavel) msg += "  👤 "+t.responsavel+"\n";
    msg += "\n";
  });

  _enviar(chatId, msg);
  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

// ============================================================
// RESUMO FROTA
// ============================================================
function _enviarResumo(chatId) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var status = { baixada:0, operando:0, reserva:0 };

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var h = String(d[i][COL_STATUS_H]||"").toUpperCase();
      if (!h) continue;
      if      (h.indexOf("BAIXA")   >= 0) status.baixada++;
      else if (h.indexOf("RESERVA") >= 0) status.reserva++;
      else                                status.operando++;
    }
  });

  var alertas = calcularAlertas(ss);
  var tarefas = lerTarefas(ss);

  var msg = "📊 *RESUMO FROTA 17ºGB*\n"
    +"_"+obterDataHoraBrasil()+"_\n――――――――――――\n"
    +"🚒 Baixadas: *"+status.baixada+"*\n"
    +"🚗 Operando: *"+status.operando+"*\n"
    +"⏸️ Reserva: *"+status.reserva+"*\n――――――――――――\n"
    +"⏰ Alertas: *"+alertas.total+"*\n"
    +"📋 Tarefas pendentes: *"+tarefas.pendente+"*\n";

  _limparEstado(chatId);
  _enviar(chatId, msg);
  Utilities.sleep(300);
  _enviarMenu(chatId,"");
}

// ============================================================
// ALERTAS ABASTECIMENTO
// ============================================================
function _enviarAbast(chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = calcularAlertasAbastecimento(ss);

  var msg = "⛽ *ABASTECIMENTO*\n――――――――――――\n"
    +"🚒 Última VTR: *"+r.ultimaVtr+"* — "+r.ultimaData+"\n"
    +"📋 Hoje: *"+r.totalDia+"*\n";

  if (r.total > 0) {
    msg += "🚨 Sem abast. há +"+ABAST_DIAS_ALERTA+"d: *"+r.total+"*\n――――――――――――\n";
    r.atrasados.slice(0,10).forEach(function(a) {
      msg += "  • "+a.prefixo
        +(a.diasDesde > 0 ? " ("+a.diasDesde+"d)" : " (sem registro)")+"\n";
    });
    if (r.atrasados.length > 10)
      msg += "  _...e mais "+(r.atrasados.length-10)+"_\n";
  } else {
    msg += "✅ Todas abastecidas em dia!\n";
  }

  _limparEstado(chatId);
  _enviar(chatId, msg);
  Utilities.sleep(300);
  _enviarMenu(chatId,"");
}

// ============================================================
// HELPERS — PLANILHA
// ============================================================
function _listarVtrsAba(nomeAba) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = getAba(ss, nomeAba);
    if (!aba) { console.log("❌ Aba não encontrada: "+nomeAba); return []; }
    var ul = aba.getLastRow();
    if (ul < 2) return [];
    var dados = aba.getRange(2, 1, ul-1, 3).getValues();
    var lista = [];
    dados.forEach(function(ln, i) {
      var pref  = String(ln[COL_PREFIXO] ||"").trim();
      var placa = String(ln[COL_PLACA]   ||"").trim();
      var km    = _parseKm(ln[COL_KM_ATUAL]);
      if (pref && placa) lista.push({ prefixo:pref, placa:placa, km:km, linha:i+2 });
    });
    console.log("✅ "+nomeAba+": "+lista.length+" viaturas");
    return lista;
  } catch(e) {
    console.log("❌ _listarVtrsAba: "+e.message);
    return [];
  }
}

function _gravarCampoVtr(nomeAba, linhaReal, colIdx, valor) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = getAba(ss, nomeAba);
    if (!aba) return false;
    aba.getRange(linhaReal, colIdx+1).setValue(valor);
    SpreadsheetApp.flush();
    _limparCacheVtrs();
    return true;
  } catch(e) {
    console.log("❌ _gravarCampoVtr: "+e.message);
    return false;
  }
}

function _darBaixaTarefa(ss, linhaReal) {
  try {
    var aba = getAba(ss, ABA_TAREFAS);
    if (!aba) return false;
    aba.getRange(linhaReal, 5).setValue("CONCLUIDA");
    SpreadsheetApp.flush();
    return true;
  } catch(e) { return false; }
}

function _lerTarefasCompleto(ss) {
  var r = { total:0, pendente:0, andamento:0, concluida:0, lista:[] };
  var aba = getAba(ss, ABA_TAREFAS);
  if (!aba) { console.log("❌ Aba TAREFAS não encontrada: "+ABA_TAREFAS); return r; }
  var ul = aba.getLastRow();
  if (ul < 2) return r;

  aba.getRange(2, 1, ul-1, 5).getValues().forEach(function(ln, i) {
    var prefixo = String(ln[0]||"").trim();
    var desc    = String(ln[2]||"").trim();
    var resp    = String(ln[3]||"").trim();
    var stRaw   = String(ln[4]||"").trim();
    if (!desc && !stRaw) return;
    var st = stRaw.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
    r.total++;
    if      (st === "PENDENTE")          r.pendente++;
    else if (st.indexOf("ANDAM") >= 0)   r.andamento++;
    else if (st.indexOf("CONCLU") >= 0)  r.concluida++;
    else if (stRaw)                      r.pendente++;
    r.lista.push({
      linha:       i+2,
      prefixo:     prefixo,
      descricao:   desc || "(sem descrição)",
      responsavel: resp,
      status:      st || "PENDENTE"
    });
  });
  return r;
}

function _inserirNovaTarefa(ss, prefixo, descricao, responsavel) {
  try {
    var aba = getAba(ss, ABA_TAREFAS);
    if (!aba) return false;
    var proxLinha = aba.getLastRow()+1;
    aba.getRange(proxLinha, 1).setValue(prefixo     || "");
    aba.getRange(proxLinha, 2).setValue("");
    aba.getRange(proxLinha, 3).setValue(descricao   || "");
    aba.getRange(proxLinha, 4).setValue(responsavel || "");
    aba.getRange(proxLinha, 5).setValue("PENDENTE");
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    console.log("❌ _inserirNovaTarefa: "+e.message);
    return false;
  }
}

// ============================================================
// HELPERS — TELEGRAM
// ============================================================
function _enviar(chatId, texto) {
  var url = "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/sendMessage";
  try {
    var resp = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify({ chat_id:chatId, text:texto, parse_mode:"Markdown" }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      UrlFetchApp.fetch(url, {
        method:"post", contentType:"application/json",
        payload: JSON.stringify({ chat_id:chatId, text:texto.replace(/\*/g,"").replace(/_/g,"") }),
        muteHttpExceptions: true
      });
    }
  } catch(e) { console.log("❌ _enviar: "+e.message); }
}

function _enviarTeclado(chatId, texto, teclado) {
  var url = "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/sendMessage";
  try {
    var resp = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify({
        chat_id:      chatId.toString(),
        text:         texto,
        parse_mode:   "Markdown",
        reply_markup: teclado
      }),
      muteHttpExceptions: true
    });
    console.log("_enviarTeclado: "+resp.getContentText());
  } catch(e) { console.log("    _enviarTeclado: "+e.message); }
}

// ============================================================
// PARSE DE DATA
// ============================================================
function _parseDataBot(texto) {
  if (!texto) return null;
  var s = texto.trim().split(" ")[0];
  var p = s.split("/");
  if (p.length !== 3) return null;
  var dd = parseInt(p[0],10);
  var mm = parseInt(p[1],10);
  var aa = parseInt(p[2],10);
  if (isNaN(dd)||isNaN(mm)||isNaN(aa)) return null;
  if (aa < 100) aa += 2000;
  if (dd<1||dd>31||mm<1||mm>12||aa<2020||aa>2099) return null;
  return new Date(aa, mm-1, dd);
}

// ============================================================
// UTILITÁRIOS WEBHOOK
// ============================================================
function configurarWebhook() {
  var urlExec = ScriptApp.getService().getUrl().replace("/dev","/exec");
  var resp    = UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/setWebhook?url="+urlExec,
    { muteHttpExceptions:true });
  var json = JSON.parse(resp.getContentText());
  console.log(JSON.stringify(json));
}

function verificarWebhook() {
  var resp = UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/getWebhookInfo",
    { muteHttpExceptions:true });
  console.log(resp.getContentText());
}

function limparFilaUpdates() {
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteWebhook?drop_pending_updates=true",
    { muteHttpExceptions:true });
  console.log("✅ Fila limpa");
  Utilities.sleep(1000);
  configurarWebhook();
  Utilities.sleep(500);
  verificarWebhook();
}

function limparUpdatesAntigos() {
  var prop  = PropertiesService.getScriptProperties();
  var todas = prop.getProperties();
  var cnt   = 0;
  Object.keys(todas).forEach(function(k) {
    if (k.indexOf("upd_") === 0) { prop.deleteProperty(k); cnt++; }
  });
  console.log("🧹 " + cnt + " update IDs removidos");
}

function limparEstadoChat() {
  PropertiesService.getScriptProperties().deleteProperty("estado_"+CHAT_ID);
  console.log("✅ Estado limpo!");
}

function diagnosticarEstado() {
  var prop  = PropertiesService.getScriptProperties();
  var todas = prop.getProperties();
  var cnt   = 0;
  console.log("=== ESTADO ===");
  Object.keys(todas).forEach(function(k) {
    if (k.indexOf("upd_") === 0) { cnt++; return; }
    console.log(k + " = " + todas[k]);
  });
  console.log("Updates em cache: " + cnt);
}

// ============================================================
// TESTES
// ============================================================
function testarDoPostLocal() {
  limparEstadoChat();
  doPost({postData:{contents:JSON.stringify({
    update_id: Math.floor(Math.random()*999999),
    message:{message_id:1,from:{id:123456,first_name:"Teste"},
      chat:{id:-5205191691,title:"MOTOMEC 17GB",type:"group"},
      date:Math.floor(Date.now()/1000),text:"/start"}
  })}});
}
// ============================================================
// ESTADO DA CONVERSA
// ============================================================
function _getEstado(chatId) {
  var raw = PropertiesService.getScriptProperties()
              .getProperty("estado_"+chatId);
  return raw ? JSON.parse(raw) : { etapa: "MENU" };
}
function _setEstado(chatId, estado) {
  PropertiesService.getScriptProperties()
    .setProperty("estado_"+chatId, JSON.stringify(estado));
}
function _limparEstado(chatId) {
  PropertiesService.getScriptProperties()
    .deleteProperty("estado_"+chatId);
}

// ============================================================
// WEBHOOK — apenas UMA definição
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var update   = JSON.parse(e.postData.contents);
    var updateId = String(update.update_id || "");
    var prop     = PropertiesService.getScriptProperties();

    // Anti-duplicata
    if (updateId && prop.getProperty("upd_" + updateId)) {
      console.log("⚠️ Update " + updateId + " já processado — ignorado");
      return ContentService.createTextOutput("OK");
    }

    // Lock máx 8 segundos
    if (!lock.tryLock(8000)) {
      console.log("⚠️ Lock ocupado — ignorando " + updateId);
      return ContentService.createTextOutput("OK");
    }

    // Dupla verificação após lock
    if (updateId && prop.getProperty("upd_" + updateId)) {
      console.log("⚠️ Update " + updateId + " já processado (pós-lock)");
      return ContentService.createTextOutput("OK");
    }

    if (updateId) prop.setProperty("upd_" + updateId, "1");
    _processarUpdate(update);

  } catch(err) {
    console.log("❌ doPost: " + err.message);
  } finally {
    try { lock.releaseLock(); } catch(e2) {}
  }
  return ContentService.createTextOutput("OK");
}

function doGet(e) {
  return ContentService.createTextOutput("✅ Bot FROTA 17ºGB online!");
}

// ============================================================
// PROCESSAMENTO
// ============================================================
function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName;

  if (callback) {
    chatId   = callback.message.chat.id.toString();
    texto    = callback.data;
    fromName = callback.from.first_name || "Auxiliar";
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId   = update.message.chat.id.toString();
    texto    = update.message.text || "";
    fromName = update.message.from.first_name || "Auxiliar";
  } else {
    return;
  }

  // Verifica chat autorizado
  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName);
}

function _responderCallback(callbackId) {
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/answerCallbackQuery",
    { method:"post", contentType:"application/json",
      payload: JSON.stringify({callback_query_id: callbackId}),
      muteHttpExceptions: true }
  );
}

// ============================================================
// ROTEADOR
// ============================================================
function _rotear(chatId, texto, fromName) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName);
    return;
  }

  // Qualquer etapa pode voltar ao menu
  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName);
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName);
  }
}

// ============================================================
// MENU PRINCIPAL
// ============================================================
function _enviarMenu(chatId, nome, messageIdParaDeletar) {
  _setEstado(chatId, { etapa:"MENU" });
  
  // Deleta mensagem anterior se informada
  if (messageIdParaDeletar) {
    try {
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteMessage",
        { method:"post", contentType:"application/json",
          payload: JSON.stringify({ chat_id:chatId, message_id:messageIdParaDeletar }),
          muteHttpExceptions:true }
      );
    } catch(e) {}
  }

  var msg = "👋 Olá, *"+(nome||"Auxiliar")+"*!\n\n"
    + "📋 *FROTA 17ºGB — Menu Principal*\n\n"
    + "O que deseja fazer?";
  var teclado = { inline_keyboard: [
    [{ text:"🚒 Atualizar Viatura",  callback_data:"MENU_VTR"    }],
    [{ text:"📋 Gerenciar Tarefas",  callback_data:"MENU_TAR"    }],
    [{ text:"📊 Ver Resumo Frota",   callback_data:"MENU_RESUMO" }],
    [{ text:"⛽ Ver Alertas Abast.", callback_data:"MENU_ABAST"  }]
  ]};
  _enviarTeclado(chatId, msg, teclado);
}

function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName, messageId;

  if (callback) {
    chatId    = callback.message.chat.id.toString();
    texto     = callback.data;
    fromName  = callback.from.first_name || "Auxiliar";
    messageId = callback.message.message_id; // ← captura ID da mensagem clicada
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId    = update.message.chat.id.toString();
    texto     = update.message.text || "";
    fromName  = update.message.from.first_name || "Auxiliar";
    messageId = null;
  } else {
    return;
  }

  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName, messageId);
}

function _rotear(chatId, texto, fromName, messageId) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName, messageId);
  }
}

function _tratarMenu(chatId, texto, fromName) {
  switch(texto) {
    case "MENU_VTR":    _iniciarVtr(chatId);     break;
    case "MENU_TAR":    _iniciarTarefas(chatId); break;
    case "MENU_RESUMO": _enviarResumo(chatId);   break;
    case "MENU_ABAST":  _enviarAbast(chatId);    break;
    default: _enviarMenu(chatId, fromName);
  }
}

// ============================================================
// FLUXO — ATUALIZAR VIATURA
// ============================================================
function _iniciarVtr(chatId) {
  _setEstado(chatId, { etapa:"VTR_ESCOLHER_ABA" });
  var teclado = { inline_keyboard: [
    [{ text:"1️⃣ 1SGB", callback_data:"ABA_1SGB" },
     { text:"2️⃣ 2SGB", callback_data:"ABA_2SGB" }],
    [{ text:"🔙 Voltar", callback_data:"VOLTAR_MENU" }]
  ]};
  _enviarTeclado(chatId, "🚒 Qual subgrupamento?", teclado);
}

function _tratarEscolherAba(chatId, texto, estado) {
  if (texto !== "ABA_1SGB" && texto !== "ABA_2SGB") {
    _iniciarVtr(chatId); return;
  }

  var nomeAba = texto === "ABA_1SGB" ? ABA_1SGB : ABA_2SGB;
  var vtrs    = _listarVtrsAbaCached(nomeAba);

  if (!vtrs || !vtrs.length) {
    _enviar(chatId, "⚠️ Nenhuma viatura encontrada em: " + nomeAba);
    return;
  }

  _setEstado(chatId, { etapa:"VTR_ESCOLHER_VTR", aba:nomeAba, abaKey:texto, vtrs:vtrs });

  var linhas = [];
  for (var i = 0; i < vtrs.length; i += 2) {
    var linha = [{ text:vtrs[i].prefixo, callback_data:"VTR_"+i }];
    if (vtrs[i+1]) linha.push({ text:vtrs[i+1].prefixo, callback_data:"VTR_"+(i+1) });
    linhas.push(linha);
  }
  linhas.push([{ text:"🔙 Voltar", callback_data:"VOLTAR_MENU" }]);

  _enviarTeclado(chatId,
    "🚒 *"+nomeAba+"* — "+vtrs.length+" viaturas\nSelecione:",
    { inline_keyboard: linhas });
}

function _tratarEscolherVtr(chatId, texto, estado) {
  if (texto === "VOLTAR_ABA") {
    _iniciarVtr(chatId); return;
  }

  if (texto.indexOf("VTR_") !== 0) {
    _iniciarVtr(chatId); return;
  }

  var idx = parseInt(texto.replace("VTR_",""), 10);
  if (isNaN(idx) || !estado.vtrs || !estado.vtrs[idx]) {
    _enviar(chatId, "⚠️ Viatura inválida. Selecione novamente.");
    Utilities.sleep(300);
    // Reenvia lista de viaturas
    _tratarEscolherAba(chatId, estado.abaKey, { etapa:"VTR_ESCOLHER_ABA" });
    return;
  }

  var vtr = estado.vtrs[idx];
  _setEstado(chatId, {
    etapa:  "VTR_ESCOLHER_CAMPO",
    aba:    estado.aba,
    abaKey: estado.abaKey,
    vtrIdx: idx,
    vtr:    vtr,
    vtrs:   estado.vtrs
  });

  var info = "🚒 *"+vtr.prefixo+"* — "+vtr.placa+"\n"
    + "📍 KM atual: *"+vtr.km.toLocaleString("pt-BR")+"*\n\n"
    + "Qual campo deseja atualizar?";

  var teclado = { inline_keyboard: [
    [{ text:"💧 Próx. Troca Óleo (KM)",   callback_data:"CAMPO_OleoKm"    },
     { text:"📅 Próx. Troca Óleo (Data)", callback_data:"CAMPO_OleoData"  }],
    [{ text:"🛑 Revisão Freio (KM)",      callback_data:"CAMPO_FreioKm"   },
     { text:"🔋 Garantia Bateria (Data)", callback_data:"CAMPO_Bateria"   }],
    [{ text:"🚿 Data Lavagem",            callback_data:"CAMPO_Lavagem"   },
     { text:"🔵 Garantia VTR (Data)",     callback_data:"CAMPO_Garantia"  }],
    [{ text:"🟡 Pneu: KM Próx. Troca",   callback_data:"CAMPO_Pneu"      },
     { text:"⚙️ KM Troca Embreagem",     callback_data:"CAMPO_Embreagem" }],
    [{ text:"🔙 Voltar",                  callback_data:"VOLTAR_ABA"      }]
  ]};

  _enviarTeclado(chatId, info, teclado);
}

function _tratarEscolherCampo(chatId, texto, estado) {
  if (texto === "VOLTAR_VTR" || texto === "VOLTAR_ABA") {
    // Volta para lista de viaturas
    _setEstado(chatId, { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
    _tratarEscolherAba(chatId, estado.abaKey, { etapa:"VTR_ESCOLHER_ABA" });
    return;
  }

  var CAMPOS = {
    "CAMPO_OleoKm"   :{ nome:"Próx. Troca Óleo (KM)",   col:4,  tipo:"km",   ex:"Ex: 165000"     },
    "CAMPO_OleoData" :{ nome:"Próx. Troca Óleo (Data)", col:5,  tipo:"data", ex:"Ex: 26/06/2026" },
    "CAMPO_FreioKm"  :{ nome:"Revisão Freio (KM)",      col:6,  tipo:"km",   ex:"Ex: 170000"     },
    "CAMPO_Bateria"  :{ nome:"Data Venc. Bateria",      col:7,  tipo:"data", ex:"Ex: 26/10/2026" },
    "CAMPO_Lavagem"  :{ nome:"Data Última Lavagem",     col:12, tipo:"data", ex:"Ex: 20/03/2026" },
    "CAMPO_Garantia" :{ nome:"VTR em Garantia (Data)",  col:3,  tipo:"data", ex:"Ex: 09/03/2027" },
    "CAMPO_Pneu"     :{ nome:"Pneus: KM Próx. Troca",  col:13, tipo:"km",   ex:"Ex: 180000"     },
    "CAMPO_Embreagem":{ nome:"KM Troca Embreagem",      col:14, tipo:"km",   ex:"Ex: 200000"     }
  };

  var campo = CAMPOS[texto];
  if (!campo) {
    _enviar(chatId, "⚠️ Campo inválido.");
    return;
  }

  _setEstado(chatId, {
    etapa:    "VTR_INSERIR_VALOR",
    aba:      estado.aba,
    abaKey:   estado.abaKey,
    vtr:      estado.vtr,
    vtrIdx:   estado.vtrIdx,
    vtrs:     estado.vtrs,
    campo:    campo,
    campoKey: texto
  });

  _enviar(chatId,
    "✏️ *"+estado.vtr.prefixo+"* — "+campo.nome+"\n\n"
    +"Digite o novo valor:\n_"+campo.ex+"_\n\n"
    +"Ou /cancelar para voltar.");
}

function _tratarInserirValor(chatId, texto, estado) {
  if (texto === "/cancelar") {
    _setEstado(chatId, {
      etapa:  "VTR_ESCOLHER_CAMPO",
      aba:    estado.aba,
      abaKey: estado.abaKey,
      vtr:    estado.vtr,
      vtrIdx: estado.vtrIdx,
      vtrs:   estado.vtrs
    });
    _tratarEscolherVtr(chatId, "VTR_"+estado.vtrIdx,
      { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
    return;
  }

  var campo = estado.campo;
  var valor, valorFmt;

  if (campo.tipo === "km") {
    var n = parseInt(texto.replace(/\D/g,""), 10);
    if (isNaN(n) || n < 1000 || n > 9999999) {
      _enviar(chatId, "⚠️ KM inválido. Digite apenas números.\n_"+campo.ex+"_");
      return;
    }
    valor    = n;
    valorFmt = n.toLocaleString("pt-BR")+" km";
  } else {
    var dt = _parseDataBot(texto);
    if (!dt) {
      _enviar(chatId, "⚠️ Data inválida. Use DD/MM/AAAA.\n_Ex: 26/06/2026_");
      return;
    }
    valor    = dt;
    valorFmt = _fmtData(dt);
  }

  var ok = _gravarCampoVtr(estado.aba, estado.vtr.linha, campo.col, valor);

  if (ok) {
    _enviar(chatId,
      "✅ *Atualizado com sucesso!*\n\n"
      +"🚒 "+estado.vtr.prefixo+" — "+estado.vtr.placa+"\n"
      +"📝 "+campo.nome+": *"+valorFmt+"*");
  } else {
    _enviar(chatId, "❌ Erro ao gravar. Tente novamente.");
  }

  Utilities.sleep(300);
  _setEstado(chatId, {
    etapa:  "VTR_ESCOLHER_CAMPO",
    aba:    estado.aba,
    abaKey: estado.abaKey,
    vtr:    estado.vtr,
    vtrIdx: estado.vtrIdx,
    vtrs:   estado.vtrs
  });
  _tratarEscolherVtr(chatId, "VTR_"+estado.vtrIdx,
    { etapa:"VTR_ESCOLHER_VTR", aba:estado.aba, abaKey:estado.abaKey, vtrs:estado.vtrs });
}

// ============================================================
// FLUXO — TAREFAS
// ============================================================
function _iniciarTarefas(chatId) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas = _lerTarefasCompleto(ss);

  var msg = "📋 *TAREFAS — "+tarefas.total+" registros*\n\n";
  if (tarefas.pendente  > 0) msg += "🔴 PENDENTE: "     +tarefas.pendente+"\n";
  if (tarefas.andamento > 0) msg += "🟡 EM ANDAMENTO: " +tarefas.andamento+"\n";
  if (tarefas.concluida > 0) msg += "✅ CONCLUÍDA: "    +tarefas.concluida+"\n";

  _setEstado(chatId, { etapa:"TAR_MENU" });

  var teclado = { inline_keyboard: [
    [{ text:"✅ Dar Baixa em Tarefa", callback_data:"TAR_BAIXA"   }],
    [{ text:"➕ Nova Tarefa",         callback_data:"TAR_NOVA"    }],
    [{ text:"📋 Listar Pendentes",    callback_data:"TAR_LISTAR"  }],
    [{ text:"🔙 Menu Principal",      callback_data:"VOLTAR_MENU" }]
  ]};

  _enviarTeclado(chatId, msg, teclado);
}

function _tratarMenuTarefas(chatId, texto, estado) {
  switch(texto) {
    case "TAR_BAIXA":  _iniciarBaixaTarefa(chatId);     break;
    case "TAR_NOVA":   _iniciarNovaTarefa(chatId);      break;
    case "TAR_LISTAR": _listarTarefasPendentes(chatId); break;
    default: _iniciarTarefas(chatId);
  }
}

function _iniciarBaixaTarefa(chatId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas   = _lerTarefasCompleto(ss);
  var pendentes = tarefas.lista.filter(function(t) {
    return t.status !== "CONCLUIDA";
  });

  if (!pendentes.length) {
    _enviar(chatId,"✅ Não há tarefas pendentes!");
    Utilities.sleep(300);
    _iniciarTarefas(chatId);
    return;
  }

  _setEstado(chatId, { etapa:"TAR_ESCOLHER_BAIXA", pendentes:pendentes });

  var linhas = [];
  var max = Math.min(pendentes.length, 20);
  for (var i = 0; i < max; i++) {
    var t   = pendentes[i];
    var lbl = (i+1)+". "+t.descricao.substring(0,30)
              +(t.descricao.length>30?"...":"")+" ["+t.status+"]";
    linhas.push([{ text:lbl, callback_data:"BAIXA_"+i }]);
  }
  linhas.push([{ text:"🔙 Voltar", callback_data:"VOLTAR_TAR" }]);

  _enviarTeclado(chatId, "✅ *Selecione a tarefa concluída:*",
    { inline_keyboard: linhas });
}

function _tratarEscolherBaixa(chatId, texto, estado) {
  if (texto === "VOLTAR_TAR") { _iniciarTarefas(chatId); return; }

  var idx = parseInt(texto.replace("BAIXA_",""), 10);
  if (isNaN(idx) || !estado.pendentes[idx]) {
    _enviar(chatId,"⚠️ Seleção inválida."); return;
  }

  var tarefa = estado.pendentes[idx];
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ok     = _darBaixaTarefa(ss, tarefa.linha);

  _enviar(chatId, ok
    ? "✅ *Tarefa concluída!*\n\n📝 "+tarefa.descricao+"\n👤 "+tarefa.responsavel
    : "❌ Erro ao dar baixa. Tente novamente.");

  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

function _iniciarNovaTarefa(chatId) {
  _setEstado(chatId, { etapa:"TAR_NOVA_PREFIXO" });
  _enviar(chatId,
    "➕ *Nova Tarefa*\n\n"
    +"🚒 Digite o *PREFIXO* da viatura:\n"
    +"_Ex: ABS-17101_\n_(ou /cancelar para voltar)_");
}

function _tratarNovaPrefixo(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }
  _setEstado(chatId, { etapa:"TAR_NOVA_DESC", prefixo:texto });
  _enviar(chatId, "📝 Digite a *DESCRIÇÃO* da tarefa:\n_(ou /cancelar)_");
}

function _tratarNovaDesc(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }
  if (texto.length < 3) { _enviar(chatId,"⚠️ Descrição muito curta."); return; }
  _setEstado(chatId, { etapa:"TAR_NOVA_RESP", prefixo:estado.prefixo, descricao:texto });
  _enviar(chatId, "👤 Quem é o *RESPONSÁVEL*?\n_(ou /cancelar)_");
}

function _tratarNovaResp(chatId, texto, estado) {
  if (texto === "/cancelar") { _iniciarTarefas(chatId); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ok = _inserirNovaTarefa(ss, estado.prefixo, estado.descricao, texto);

  _enviar(chatId, ok
    ? "✅ *Tarefa criada!*\n\n🚒 "+estado.prefixo+"\n📝 "+estado.descricao+"\n👤 "+texto+"\n🔴 Status: PENDENTE"
    : "❌ Erro ao criar tarefa.");

  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

function _listarTarefasPendentes(chatId) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas   = _lerTarefasCompleto(ss);
  var pendentes = tarefas.lista.filter(function(t) {
    return t.status !== "CONCLUIDA";
  });

  if (!pendentes.length) {
    _enviar(chatId,"✅ Nenhuma tarefa pendente!");
    Utilities.sleep(300);
    _iniciarTarefas(chatId);
    return;
  }

  var msg = "📋 *TAREFAS PENDENTES ("+pendentes.length+")*\n――――――――――――\n";
  pendentes.forEach(function(t, i) {
    var icone = t.status === "EM ANDAMENTO" ? "🟡" : "🔴";
    msg += icone+" *"+(i+1)+".* "+t.descricao+"\n";
    if (t.responsavel) msg += "  👤 "+t.responsavel+"\n";
    msg += "\n";
  });

  _enviar(chatId, msg);
  Utilities.sleep(300);
  _iniciarTarefas(chatId);
}

// ============================================================
// RESUMO FROTA
// ============================================================
function _enviarResumo(chatId) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var status = { baixada:0, operando:0, reserva:0 };

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var h = String(d[i][COL_STATUS_H]||"").toUpperCase();
      if (!h) continue;
      if      (h.indexOf("BAIXA")   >= 0) status.baixada++;
      else if (h.indexOf("RESERVA") >= 0) status.reserva++;
      else                                status.operando++;
    }
  });

  var alertas = calcularAlertas(ss);
  var tarefas = lerTarefas(ss);

  var msg = "📊 *RESUMO FROTA 17ºGB*\n"
    +"_"+obterDataHoraBrasil()+"_\n――――――――――――\n"
    +"🚒 Baixadas: *"+status.baixada+"*\n"
    +"🚗 Operando: *"+status.operando+"*\n"
    +"⏸️ Reserva: *"+status.reserva+"*\n――――――――――――\n"
    +"⏰ Alertas: *"+alertas.total+"*\n"
    +"📋 Tarefas pendentes: *"+tarefas.pendente+"*\n";

  _limparEstado(chatId);
  _enviar(chatId, msg);
  Utilities.sleep(300);
  _enviarMenu(chatId,"");
}

// ============================================================
// ALERTAS ABASTECIMENTO
// ============================================================
function _enviarAbast(chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r  = calcularAlertasAbastecimento(ss);

  var msg = "⛽ *ABASTECIMENTO*\n――――――――――――\n"
    +"🚒 Última VTR: *"+r.ultimaVtr+"* — "+r.ultimaData+"\n"
    +"📋 Hoje: *"+r.totalDia+"*\n";

  if (r.total > 0) {
    msg += "🚨 Sem abast. há +"+ABAST_DIAS_ALERTA+"d: *"+r.total+"*\n――――――――――――\n";
    r.atrasados.slice(0,10).forEach(function(a) {
      msg += "  • "+a.prefixo
        +(a.diasDesde > 0 ? " ("+a.diasDesde+"d)" : " (sem registro)")+"\n";
    });
    if (r.atrasados.length > 10)
      msg += "  _...e mais "+(r.atrasados.length-10)+"_\n";
  } else {
    msg += "✅ Todas abastecidas em dia!\n";
  }

  _limparEstado(chatId);
  _enviar(chatId, msg);
  Utilities.sleep(300);
  _enviarMenu(chatId,"");
}

// ============================================================
// HELPERS — PLANILHA
// ============================================================
function _listarVtrsAba(nomeAba) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = getAba(ss, nomeAba);
    if (!aba) { console.log("❌ Aba não encontrada: "+nomeAba); return []; }
    var ul = aba.getLastRow();
    if (ul < 2) return [];
    var dados = aba.getRange(2, 1, ul-1, 3).getValues();
    var lista = [];
    dados.forEach(function(ln, i) {
      var pref  = String(ln[COL_PREFIXO] ||"").trim();
      var placa = String(ln[COL_PLACA]   ||"").trim();
      var km    = _parseKm(ln[COL_KM_ATUAL]);
      if (pref && placa) lista.push({ prefixo:pref, placa:placa, km:km, linha:i+2 });
    });
    console.log("✅ "+nomeAba+": "+lista.length+" viaturas");
    return lista;
  } catch(e) {
    console.log("❌ _listarVtrsAba: "+e.message);
    return [];
  }
}

function _gravarCampoVtr(nomeAba, linhaReal, colIdx, valor) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = getAba(ss, nomeAba);
    if (!aba) return false;
    aba.getRange(linhaReal, colIdx+1).setValue(valor);
    SpreadsheetApp.flush();
    _limparCacheVtrs();
    return true;
  } catch(e) {
    console.log("❌ _gravarCampoVtr: "+e.message);
    return false;
  }
}

function _darBaixaTarefa(ss, linhaReal) {
  try {
    var aba = getAba(ss, ABA_TAREFAS);
    if (!aba) return false;
    aba.getRange(linhaReal, 5).setValue("CONCLUIDA");
    SpreadsheetApp.flush();
    return true;
  } catch(e) { return false; }
}

function _lerTarefasCompleto(ss) {
  var r = { total:0, pendente:0, andamento:0, concluida:0, lista:[] };
  var aba = getAba(ss, ABA_TAREFAS);
  if (!aba) { console.log("❌ Aba TAREFAS não encontrada: "+ABA_TAREFAS); return r; }
  var ul = aba.getLastRow();
  if (ul < 2) return r;

  aba.getRange(2, 1, ul-1, 5).getValues().forEach(function(ln, i) {
    var prefixo = String(ln[0]||"").trim();
    var desc    = String(ln[2]||"").trim();
    var resp    = String(ln[3]||"").trim();
    var stRaw   = String(ln[4]||"").trim();
    if (!desc && !stRaw) return;
    var st = stRaw.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
    r.total++;
    if      (st === "PENDENTE")          r.pendente++;
    else if (st.indexOf("ANDAM") >= 0)   r.andamento++;
    else if (st.indexOf("CONCLU") >= 0)  r.concluida++;
    else if (stRaw)                      r.pendente++;
    r.lista.push({
      linha:       i+2,
      prefixo:     prefixo,
      descricao:   desc || "(sem descrição)",
      responsavel: resp,
      status:      st || "PENDENTE"
    });
  });
  return r;
}

function _inserirNovaTarefa(ss, prefixo, descricao, responsavel) {
  try {
    var aba = getAba(ss, ABA_TAREFAS);
    if (!aba) return false;
    var proxLinha = aba.getLastRow()+1;
    aba.getRange(proxLinha, 1).setValue(prefixo     || "");
    aba.getRange(proxLinha, 2).setValue("");
    aba.getRange(proxLinha, 3).setValue(descricao   || "");
    aba.getRange(proxLinha, 4).setValue(responsavel || "");
    aba.getRange(proxLinha, 5).setValue("PENDENTE");
    SpreadsheetApp.flush();
    return true;
  } catch(e) {
    console.log("❌ _inserirNovaTarefa: "+e.message);
    return false;
  }
}

// ============================================================
// HELPERS — TELEGRAM
// ============================================================
function _enviar(chatId, texto) {
  var url = "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/sendMessage";
  try {
    var resp = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify({ chat_id:chatId, text:texto, parse_mode:"Markdown" }),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      UrlFetchApp.fetch(url, {
        method:"post", contentType:"application/json",
        payload: JSON.stringify({ chat_id:chatId, text:texto.replace(/\*/g,"").replace(/_/g,"") }),
        muteHttpExceptions: true
      });
    }
  } catch(e) { console.log("❌ _enviar: "+e.message); }
}

function _enviarTeclado(chatId, texto, teclado) {
  var url = "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/sendMessage";
  try {
    var resp = UrlFetchApp.fetch(url, {
      method:"post", contentType:"application/json",
      payload: JSON.stringify({
        chat_id:      chatId.toString(),
        text:         texto,
        parse_mode:   "Markdown",
        reply_markup: teclado
      }),
      muteHttpExceptions: true
    });
    console.log("_enviarTeclado: "+resp.getContentText());
  } catch(e) { console.log("    _enviarTeclado: "+e.message); }
}

// ============================================================
// PARSE DE DATA
// ============================================================
function _parseDataBot(texto) {
  if (!texto) return null;
  var s = texto.trim().split(" ")[0];
  var p = s.split("/");
  if (p.length !== 3) return null;
  var dd = parseInt(p[0],10);
  var mm = parseInt(p[1],10);
  var aa = parseInt(p[2],10);
  if (isNaN(dd)||isNaN(mm)||isNaN(aa)) return null;
  if (aa < 100) aa += 2000;
  if (dd<1||dd>31||mm<1||mm>12||aa<2020||aa>2099) return null;
  return new Date(aa, mm-1, dd);
}

// ============================================================
// UTILITÁRIOS WEBHOOK
// ============================================================
function configurarWebhook() {
  var urlExec = ScriptApp.getService().getUrl().replace("/dev","/exec");
  var resp    = UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/setWebhook?url="+urlExec,
    { muteHttpExceptions:true });
  var json = JSON.parse(resp.getContentText());
  console.log(JSON.stringify(json));
}

function verificarWebhook() {
  var resp = UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/getWebhookInfo",
    { muteHttpExceptions:true });
  console.log(resp.getContentText());
}

function limparFilaUpdates() {
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteWebhook?drop_pending_updates=true",
    { muteHttpExceptions:true });
  console.log("✅ Fila limpa");
  Utilities.sleep(1000);
  configurarWebhook();
  Utilities.sleep(500);
  verificarWebhook();
}

function limparUpdatesAntigos() {
  var prop  = PropertiesService.getScriptProperties();
  var todas = prop.getProperties();
  var cnt   = 0;
  Object.keys(todas).forEach(function(k) {
    if (k.indexOf("upd_") === 0) { prop.deleteProperty(k); cnt++; }
  });
  console.log("🧹 " + cnt + " update IDs removidos");
}

function limparEstadoChat() {
  PropertiesService.getScriptProperties().deleteProperty("estado_"+CHAT_ID);
  console.log("✅ Estado limpo!");
}

function diagnosticarEstado() {
  var prop  = PropertiesService.getScriptProperties();
  var todas = prop.getProperties();
  var cnt   = 0;
  console.log("=== ESTADO ===");
  Object.keys(todas).forEach(function(k) {
    if (k.indexOf("upd_") === 0) { cnt++; return; }
    console.log(k + " = " + todas[k]);
  });
  console.log("Updates em cache: " + cnt);
}

// ============================================================
// TESTES
// ============================================================
function testarDoPostLocal() {
  limparEstadoChat();
  doPost({postData:{contents:JSON.stringify({
    update_id: Math.floor(Math.random()*999999),
    message:{message_id:1,from:{id:123456,first_name:"Teste"},
      chat:{id:-5205191691,title:"MOTOMEC 17GB",type:"group"},
      date:Math.floor(Date.now()/1000),text:"/start"}
  })}});
}
function verificarFuncoes() {
  var funcoes = [
    "_getEstado", "_setEstado", "_limparEstado",
    "doPost", "doGet", "_processarUpdate", "_responderCallback",
    "_rotear", "_enviarMenu", "_tratarMenu",
    "_iniciarVtr", "_tratarEscolherAba", "_tratarEscolherVtr",
    "_tratarEscolherCampo", "_tratarInserirValor",
    "_iniciarTarefas", "_tratarMenuTarefas", "_iniciarBaixaTarefa",
    "_tratarEscolherBaixa", "_iniciarNovaTarefa",
    "_tratarNovaPrefixo", "_tratarNovaDesc", "_tratarNovaResp",
    "_listarTarefasPendentes", "_enviarResumo", "_enviarAbast",
    "_listarVtrsAba", "_listarVtrsAbaCached", "_limparCacheVtrs",
    "_gravarCampoVtr", "_darBaixaTarefa",
    "_lerTarefasCompleto", "_inserirNovaTarefa",
    "_enviar", "_enviarTeclado", "_parseDataBot",
    "configurarWebhook", "verificarWebhook",
    "limparFilaUpdates", "limparUpdatesAntigos",
    "limparEstadoChat", "diagnosticarEstado"
  ];

  funcoes.forEach(function(f) {
    try {
      var existe = (typeof eval(f) === "function");
      console.log((existe ? "✅" : "❌") + " " + f);
    } catch(e) {
      console.log("❌ " + f + " — " + e.message);
    }
  });
}
function deployCompleto() {
  // Limpa estado e fila
  limparEstadoChat();
  limparUpdatesAntigos();
  limparFilaUpdates(); // já reconfigura webhook automaticamente
}
function corrigirTudo() {
  // 1. Aponta webhook para URL CORRETA (novo deploy)
  var urlCorreta = "https://script.google.com/macros/s/AKfycbyqxzLNeIQC9bkOWqgMMBcyCebWW5imvoARi_iVjUE1FhexM9259UBbVwXxV5lSkg/exec";
  
  UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + TOKEN_TELEGRAM 
    + "/deleteWebhook?drop_pending_updates=true",
    { muteHttpExceptions: true }
  );
  Utilities.sleep(1000);
  
  var r = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + TOKEN_TELEGRAM 
    + "/setWebhook?url=" + urlCorreta + "&drop_pending_updates=true",
    { muteHttpExceptions: true }
  );
  console.log("Webhook: " + r.getContentText());
  
  // 2. Recria triggers de alerta
  // Remove todos os triggers existentes primeiro
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  
  // Cria trigger a cada 5 minutos para alertas
  ScriptApp.newTrigger("enviarAlertasAutomaticos")
    .timeBased()
    .everyMinutes(5)
    .create();
  
  // Cria trigger diário para limpar cache de updates
  ScriptApp.newTrigger("limparUpdatesAntigos")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  
  console.log("✅ Triggers recriados: " + ScriptApp.getProjectTriggers().length);
  
  // 3. Verifica webhook
  Utilities.sleep(500);
  var info = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" + TOKEN_TELEGRAM + "/getWebhookInfo",
    { muteHttpExceptions: true }
  );
  console.log("Status: " + info.getContentText());
}
function testarVelocidade() {
  var inicio = new Date();
  
  // Simula o que doPost faz
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vtrs = _listarVtrsAbaCached(ABA_1SGB);
  
  var fim = new Date();
  console.log("⏱️ Tempo _listarVtrsAbaCached: " + (fim-inicio) + "ms");
  console.log("📊 Viaturas encontradas: " + vtrs.length);
  
  // Testa velocidade de PropertiesService
  var t2 = new Date();
  PropertiesService.getScriptProperties().setProperty("teste","1");
  PropertiesService.getScriptProperties().getProperty("teste");
  PropertiesService.getScriptProperties().deleteProperty("teste");
  console.log("⏱️ Tempo PropertiesService: " + (new Date()-t2) + "ms");
  
  // Testa velocidade total estimada
  console.log("⏱️ Tempo total estimado doPost: " + (new Date()-inicio) + "ms");
  console.log(fim-inicio > 5000 ? "🔴 LENTO — Telegram vai reenviar!" : "🟢 OK — dentro do limite");
}
function ativarLogDetalhado() {
  // Substitui doPost temporariamente com log completo
  // Execute /start no grupo APÓS rodar esta função
  // Depois rode verLogUpdates() para ver o que chegou
  
  PropertiesService.getScriptProperties().setProperty("LOG_UPDATES", "[]");
  console.log("✅ Log ativado — agora envie /start no grupo");
}

function verLogUpdates() {
  var raw = PropertiesService.getScriptProperties().getProperty("LOG_UPDATES");
  var lista = raw ? JSON.parse(raw) : [];
  console.log("Total de chamadas doPost: " + lista.length);
  lista.forEach(function(item, i) {
    console.log("--- Chamada " + (i+1) + " ---");
    console.log("update_id: " + item.uid);
    console.log("data: " + item.data);
    console.log("hora: " + item.hora);
    console.log("ignorado: " + item.ignorado);
  });
}
function _enviarMenu(chatId, nome, messageIdParaDeletar) {
  _setEstado(chatId, { etapa:"MENU" });
  
  // Deleta mensagem anterior se informada
  if (messageIdParaDeletar) {
    try {
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteMessage",
        { method:"post", contentType:"application/json",
          payload: JSON.stringify({ chat_id:chatId, message_id:messageIdParaDeletar }),
          muteHttpExceptions:true }
      );
    } catch(e) {}
  }

  var msg = "👋 Olá, *"+(nome||"Auxiliar")+"*!\n\n"
    + "📋 *FROTA 17ºGB — Menu Principal*\n\n"
    + "O que deseja fazer?";
  var teclado = { inline_keyboard: [
    [{ text:"🚒 Atualizar Viatura",  callback_data:"MENU_VTR"    }],
    [{ text:"📋 Gerenciar Tarefas",  callback_data:"MENU_TAR"    }],
    [{ text:"📊 Ver Resumo Frota",   callback_data:"MENU_RESUMO" }],
    [{ text:"⛽ Ver Alertas Abast.", callback_data:"MENU_ABAST"  }]
  ]};
  _enviarTeclado(chatId, msg, teclado);
}

function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName, messageId;

  if (callback) {
    chatId    = callback.message.chat.id.toString();
    texto     = callback.data;
    fromName  = callback.from.first_name || "Auxiliar";
    messageId = callback.message.message_id; // ← captura ID da mensagem clicada
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId    = update.message.chat.id.toString();
    texto     = update.message.text || "";
    fromName  = update.message.from.first_name || "Auxiliar";
    messageId = null;
  } else {
    return;
  }

  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName, messageId);
}

function _rotear(chatId, texto, fromName, messageId) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName, messageId);
  }
}
function _enviarMenu(chatId, nome, messageIdParaDeletar) {
  _setEstado(chatId, { etapa:"MENU" });
  
  // Deleta mensagem anterior se informada
  if (messageIdParaDeletar) {
    try {
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteMessage",
        { method:"post", contentType:"application/json",
          payload: JSON.stringify({ chat_id:chatId, message_id:messageIdParaDeletar }),
          muteHttpExceptions:true }
      );
    } catch(e) {}
  }

  var msg = "👋 Olá, *"+(nome||"Auxiliar")+"*!\n\n"
    + "📋 *FROTA 17ºGB — Menu Principal*\n\n"
    + "O que deseja fazer?";
  var teclado = { inline_keyboard: [
    [{ text:"🚒 Atualizar Viatura",  callback_data:"MENU_VTR"    }],
    [{ text:"📋 Gerenciar Tarefas",  callback_data:"MENU_TAR"    }],
    [{ text:"📊 Ver Resumo Frota",   callback_data:"MENU_RESUMO" }],
    [{ text:"⛽ Ver Alertas Abast.", callback_data:"MENU_ABAST"  }]
  ]};
  _enviarTeclado(chatId, msg, teclado);
}

function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName, messageId;

  if (callback) {
    chatId    = callback.message.chat.id.toString();
    texto     = callback.data;
    fromName  = callback.from.first_name || "Auxiliar";
    messageId = callback.message.message_id; // ← captura ID da mensagem clicada
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId    = update.message.chat.id.toString();
    texto     = update.message.text || "";
    fromName  = update.message.from.first_name || "Auxiliar";
    messageId = null;
  } else {
    return;
  }

  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName, messageId);
}

function _rotear(chatId, texto, fromName, messageId) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);function _enviarMenu(chatId, nome, messageIdParaDeletar) {
  _setEstado(chatId, { etapa:"MENU" });
  
  // Deleta mensagem anterior se informada
  if (messageIdParaDeletar) {
    try {
      UrlFetchApp.fetch(
        "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/deleteMessage",
        { method:"post", contentType:"application/json",
          payload: JSON.stringify({ chat_id:chatId, message_id:messageIdParaDeletar }),
          muteHttpExceptions:true }
      );
    } catch(e) {}
  }

  var msg = "👋 Olá, *"+(nome||"Auxiliar")+"*!\n\n"
    + "📋 *FROTA 17ºGB — Menu Principal*\n\n"
    + "O que deseja fazer?";
  var teclado = { inline_keyboard: [
    [{ text:"🚒 Atualizar Viatura",  callback_data:"MENU_VTR"    }],
    [{ text:"📋 Gerenciar Tarefas",  callback_data:"MENU_TAR"    }],
    [{ text:"📊 Ver Resumo Frota",   callback_data:"MENU_RESUMO" }],
    [{ text:"⛽ Ver Alertas Abast.", callback_data:"MENU_ABAST"  }]
  ]};
  _enviarTeclado(chatId, msg, teclado);
}

function _processarUpdate(update) {
  var callback = update.callback_query || null;
  var chatId, texto, fromName, messageId;

  if (callback) {
    chatId    = callback.message.chat.id.toString();
    texto     = callback.data;
    fromName  = callback.from.first_name || "Auxiliar";
    messageId = callback.message.message_id; // ← captura ID da mensagem clicada
    _responderCallback(callback.id);
  } else if (update.message) {
    chatId    = update.message.chat.id.toString();
    texto     = update.message.text || "";
    fromName  = update.message.from.first_name || "Auxiliar";
    messageId = null;
  } else {
    return;
  }

  var chatIdNorm = chatId.replace("-100","").replace("-","");
  var chatIdRef  = CHAT_ID.toString().replace("-100","").replace("-","");
  if (chatIdNorm !== chatIdRef) {
    _enviar(chatId, "⛔ Acesso não autorizado.");
    return;
  }

  _rotear(chatId, texto.trim(), fromName, messageId);
}

function _rotear(chatId, texto, fromName, messageId) {
  var estado = _getEstado(chatId);

  if (texto === "/start" || texto === "/menu" || texto === "🏠 Menu Principal") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  if (texto === "VOLTAR_MENU") {
    _limparEstado(chatId);
    _enviarMenu(chatId, fromName, messageId);
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName, messageId);
  }
}
    return;
  }

  switch (estado.etapa) {
    case "MENU":
    case undefined:
      _tratarMenu(chatId, texto, fromName); break;

    case "VTR_ESCOLHER_ABA":   _tratarEscolherAba(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_VTR":   _tratarEscolherVtr(chatId, texto, estado);   break;
    case "VTR_ESCOLHER_CAMPO": _tratarEscolherCampo(chatId, texto, estado); break;
    case "VTR_INSERIR_VALOR":  _tratarInserirValor(chatId, texto, estado);  break;

    case "TAR_MENU":           _tratarMenuTarefas(chatId, texto, estado);   break;
    case "TAR_ESCOLHER_BAIXA": _tratarEscolherBaixa(chatId, texto, estado); break;
    case "TAR_NOVA_PREFIXO":   _tratarNovaPrefixo(chatId, texto, estado);   break;
    case "TAR_NOVA_DESC":      _tratarNovaDesc(chatId, texto, estado);      break;
    case "TAR_NOVA_RESP":      _tratarNovaResp(chatId, texto, estado);      break;

    default:
      _limparEstado(chatId);
      _enviarMenu(chatId, fromName, messageId);
  }
}