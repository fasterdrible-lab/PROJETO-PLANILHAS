/**
 * ╔════════════════════════════════════════════════════════╗
 * ║  CONTROLE DE FROTA 17ºGB — v9.1                       ║
 * ╚════════════════════════════════════════════════════════╝
 */

// ============================================================
// CONFIGURAÇÕES
// ============================================================
var TOKEN_TELEGRAM          = "8030450332:AAEkyiVJbGuvf2w4nBawa_SYaZWoU2WpKZs";
var CHAT_ID                 = "-5205191691";
var ID_PLANILHA_OPERACIONAL = "1Qc6zyHENWkhKiLxVWv6OIrSar_tCZP8WzyKoOjUsO6M";
var ID_CBM                  = "12PWGLdjW8RGr0vO02Hp_VpDZmG9pX5imSreZG19OTUk";

// ============================================================
// ABAS
// ============================================================
var ABA_KM_VTR  = "KM SEMANAL";
var ABA_FCD     = "FCD";
var ABA_ABAST   = "ABAST. VTR";
var ABA_1SGB    = "1SGB";
var ABA_2SGB    = "2SGB";
var ABA_FROTA   = "FROTA";
var ABA_GASTOS  = "GASTOS";
var ABA_RIV     = "RIV_2026";
var ABA_TAREFAS = "TAREFAS";

// ============================================================
// FIPE
// ============================================================
var TOKEN_FIPE      = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VySWQiOiJmMzRlZmYzZC1jNjNlLTQzOTAtOTFiMC00OGE2OGRmNzc0NmYiLCJlbWFpbCI6ImZhc3RlcmRyaWJsZUBnbWFpbC5jb20iLCJpYXQiOjE3NzIwMjQzMTZ9.K7CmjXJQ6qKs3NPmV_JsA82JITVrfduVO6N77dNOcpk";
var BASE_URL        = "https://fipe.parallelum.com.br/api/v2";
var TIPOS_API       = ["cars","motorcycles","trucks"];
var RIV_COL_PREFIXO = 0;
var RIV_COL_DATA    = 2;
var RIV_COL_VALOR   = 6;

// ============================================================
// COLUNAS — 1SGB / 2SGB (base 0)
// ============================================================
var COL_PREFIXO  = 0;
var COL_PLACA    = 1;
var COL_KM_ATUAL = 2;
var COL_STATUS_H = 7;
var COL_L        = 11;
var COL_M        = 12;
var COL_N        = 13;
var COL_O        = 14;
var COL_P        = 15;

// ============================================================
// COLUNAS — ABAST. VTR (base 0)
// A=0 Carimbo | B=1 Email | C=2 Posto | D=3 Nome | E=4 Unidade
// F=5 Prefixo | G=6 Placa | H=7 Data  | I=8 KM
// ============================================================
var ABAST_COL_CARIMBO = 0;
var ABAST_COL_PREFIXO = 5;
var ABAST_COL_PLACA   = 6;
var ABAST_COL_DATA    = 7;
var ABAST_COL_KM      = 8;

// ============================================================
// LIMIARES
// ============================================================
var LAVAGEM_DIAS      = 15;
var ALERTA_DIAS_AVISO = 3;
var ALERTA_KM_AVISO   = 5000;
var ABAST_DIAS_ALERTA = 30;

// ============================================================
// CONSTANTES
// ============================================================
var SEP = "--------------------\n";

var MAPA_STATUS = {
  "BAIXA(DEFEITO)"  : "BAIXADA",
  "BAIXA (RÁDIO)"   : "BAIXADA",
  "BAIXA(ACIDENTE)" : "BAIXADA",
  "PROC DESCARGA"   : "BAIXADA",
  "SUBST PREFIXO"   : "BAIXADA",
  "DISPONÍVEL"      : "OPERANDO",
  "OP ESPECIAL"     : "OPERANDO",
  "EM ATENDIMENTO"  : "OPERANDO",
  "MNT RÁPIDA"      : "OPERANDO",
  "RESERVA"         : "RESERVA"
};

// ============================================================
// UTILITÁRIOS
// ============================================================
function getAba(ss, nome) {
  if (!ss) { console.log("❌ getAba: ss indefinido para \"" + nome + "\""); return null; }
  return ss.getSheetByName(nome)
    || ss.getSheetByName(" " + nome)
    || ss.getSheetByName(nome.trim())
    || null;
}

function _norm(v) {
  if (!v) return "";
  return v.toString().trim().toUpperCase().replace(/-/g,"").replace(/\s/g,"");
}

function _parseKm(v) {
  if (!v || v === "") return 0;
  if (typeof v === "number") return (v > 0 && v < 1000000) ? Math.round(v) : 0;
  var n = parseInt(String(v).replace(/\D/g,""),10);
  return (n > 0 && n < 1000000) ? n : 0;
}

function _parseData(v) {
  if (!v || v === "") return null;

  // Objeto Date — corrige fuso horário para Brasília (GMT-3)
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    // Converte para horário de Brasília antes de extrair dia/mês/ano
    var formatter = new Intl.DateTimeFormat("pt-BR", {
      timeZone: "America/Sao_Paulo",
      year:  "numeric",
      month: "2-digit",
      day:   "2-digit"
    });
    var partes = formatter.formatToParts(v);
    var dia, mes, ano;
    partes.forEach(function(p) {
      if (p.type === "day")   dia = parseInt(p.value, 10);
      if (p.type === "month") mes = parseInt(p.value, 10);
      if (p.type === "year")  ano = parseInt(p.value, 10);
    });
    return new Date(ano, mes - 1, dia);
  }

  // Número serial Google Sheets
  if (typeof v === "number") {
    // Serial → UTC → ajusta para Brasília
    var utc = new Date(Math.round((v - 25569) * 86400000));
    var formatter2 = new Intl.DateTimeFormat("pt-BR", {
      timeZone: "America/Sao_Paulo",
      year:  "numeric",
      month: "2-digit",
      day:   "2-digit"
    });
    var partes2 = formatter2.formatToParts(utc);
    var dia2, mes2, ano2;
    partes2.forEach(function(p) {
      if (p.type === "day")   dia2 = parseInt(p.value, 10);
      if (p.type === "month") mes2 = parseInt(p.value, 10);
      if (p.type === "year")  ano2 = parseInt(p.value, 10);
    });
    return new Date(ano2, mes2 - 1, dia2);
  }

  // String DD/MM/AAAA
  if (typeof v === "string") {
    var s = v.trim().split(" ")[0];
    if (s.indexOf("/") >= 0) {
      var p  = s.split("/");
      var n0 = parseInt(p[0], 10);
      var n1 = parseInt(p[1], 10);
      var n2 = parseInt(p[2], 10);
      if (p.length === 3 && !isNaN(n0) && !isNaN(n1) && !isNaN(n2)) {
        if (n2 >= 2000 && n2 <= 2100 && n1 >= 1 && n1 <= 12 && n0 >= 1 && n0 <= 31)
          return new Date(n2, n1 - 1, n0); // DD/MM/AAAA
        if (n0 >= 2000 && n0 <= 2100 && n1 >= 1 && n1 <= 12 && n2 >= 1 && n2 <= 31)
          return new Date(n0, n1 - 1, n2); // AAAA/MM/DD
      }
    }
    if (s.indexOf("-") >= 0) {
      var p2 = s.split("-");
      if (p2.length === 3 && parseInt(p2[0],10) >= 2000)
        return new Date(parseInt(p2[0],10), parseInt(p2[1],10)-1, parseInt(p2[2],10));
    }
  }

  return null;
}

function _fmtData(d) {
  if (!d) return "";
  return String(d.getDate()).padStart(2,"0") + "/" +
         String(d.getMonth()+1).padStart(2,"0") + "/" +
         d.getFullYear();
}

function _colLetra(idx) {
  var letra = "", n = idx + 1;
  while (n > 0) {
    var r = (n-1) % 26;
    letra = String.fromCharCode(65+r) + letra;
    n = Math.floor((n-1)/26);
  }
  return letra;
}

function _prioridade(s) {
  if (!s) return 0;
  var u = s.toUpperCase();
  if (u.indexOf("BAIXA")   >= 0) return 3;
  if (u.indexOf("RESERVA") >= 0) return 2;
  return 1;
}

function _combinarStatus(sp, fonte) {
  Object.keys(fonte).forEach(function(k) {
    if (!sp[k] || _prioridade(fonte[k]) >= _prioridade(sp[k])) sp[k] = fonte[k];
  });
}

function _combinarReg(base, novo) {
  Object.keys(novo).forEach(function(pl) {
    if (!base[pl] || novo[pl].km > base[pl].km) base[pl] = novo[pl];
  });
}

// ============================================================
// DATA/HORA BRASIL
// ============================================================
function obterDataHoraBrasil() {
  var agora  = new Date();
  var opcoes = {
    timeZone:"America/Sao_Paulo",
    year:"numeric", month:"2-digit", day:"2-digit",
    hour:"2-digit", minute:"2-digit", second:"2-digit",
    hour12:false
  };
  var partes = new Intl.DateTimeFormat("pt-BR", opcoes).formatToParts(agora);
  var dia,mes,ano,h,mi,s;
  partes.forEach(function(p) {
    if (p.type==="day")    dia = p.value;
    if (p.type==="month")  mes = p.value;
    if (p.type==="year")   ano = p.value;
    if (p.type==="hour")   h   = p.value;
    if (p.type==="minute") mi  = p.value;
    if (p.type==="second") s   = p.value;
  });
  var dt = dia+"/"+mes+"/"+ano+" "+h+":"+mi+":"+s;
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CONTROLE") || ss.insertSheet("CONTROLE",0);
    aba.getRange("A1").setValue("ÚLTIMA SINCRONIZAÇÃO");
    aba.getRange("B1").setValue(dt);
    aba.getRange("A1:B1").setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  } catch(e) {}
  return dt;
}

function mostrarUltimaSincronizacao() {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONTROLE");
    SpreadsheetApp.getUi().alert("⏰ ÚLTIMA SINCRONIZAÇÃO:\n\n" +
      (aba ? aba.getRange("B1").getValue() : "Nunca"));
  } catch(e) { SpreadsheetApp.getUi().alert("Erro ao obter data."); }
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 FROTA")
    .addItem("▶️ Sincronização Completa",      "sincronizarFrota")
    .addItem("💰 Atualizar FIPE + Gastos",     "atualizarTudo")
    .addSeparator()
    .addItem("⚙️ Agendar (5 minutos)",         "agendar5min")
    .addItem("🧹 Remover Agendamentos",        "removerTriggers")
    .addSeparator()
    .addItem("📅 Testar Alertas",             "testarAlertas")
    .addItem("🔍 Diagnóstico L/M/N/O",        "diagnosticarMNO")
    .addItem("🔋 Diagnóstico Bateria",         "diagnosticarBateria")
    .addItem("🧹 Limpar Colunas N/O",         "limparColunasNO")
    .addItem("⛽ Testar Abastecimento",        "testarAlertasAbastecimento")
    .addItem("🐛 Debug Completo",             "debugCompleto")
    .addSeparator()
    .addItem("ℹ️ Última Sincronização",        "mostrarUltimaSincronizacao")
    .addToUi();
}

// ============================================================
// FUNÇÃO PRINCIPAL
// ============================================================
function sincronizarFrota() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("╔══════════════════════════════════════════╗");
  console.log("║  FROTA 17ºGB v9.1 — " + obterDataHoraBrasil());
  console.log("╚══════════════════════════════════════════╝\n");

  console.log("▶ PARTE 1: KM");
  var totalKm = sincronizarKm(ss);

  console.log("\n▶ PARTE 2: STATUS");
  var status = sincronizarStatus(ss);

  console.log("\n▶ PARTE 3: TAREFAS");
  var tarefas = lerTarefas(ss);

  console.log("\n▶ PARTE 4: ALERTAS");
  var alertas = calcularAlertas(ss);

  console.log("\n▶ PARTE 5: ABASTECIMENTO");
  var alertasAbast = calcularAlertasAbastecimento(ss);

  console.log("\n╔══════════════════════════════════════════╗");
  console.log("║  ✅ CONCLUÍDO | KM:" + totalKm
    + " STATUS:" + status.atualizados
    + " TAREFAS:" + tarefas.total
    + " ALERTAS:" + alertas.total
    + " ABAST:" + alertasAbast.total);
  console.log("╚══════════════════════════════════════════╝");

  enviarTelegram(totalKm, status, tarefas, alertas, alertasAbast);
}

// ============================================================
// PARTE 1: KM
// ============================================================
function sincronizarKm(ss) {
  var reg = {};
  _combinarReg(reg, _kmVtr(ss));
  _combinarReg(reg, _kmFcd(ss));
  _combinarReg(reg, _kmAbast(ss));
  console.log("  ✓ " + Object.keys(reg).length + " placas no registro");
  return _gravarKm(ss, reg);
}

function _kmVtr(ss) {
  var aba = getAba(ss, ABA_KM_VTR);
  if (!aba) return {};
  var dados = aba.getRange(2,1,Math.max(aba.getLastRow()-1,1),
              aba.getLastColumn()).getValues();
  var r = {};
  dados.forEach(function(l) {
    var pl = l[1] ? _norm(l[1]) : null;
    var km = _parseKm(l[4]);
    if (pl && km > 0) r[pl] = {km:km, ts:l[0]};
  });
  console.log("  ✓ " + Object.keys(r).length + " de " + ABA_KM_VTR);
  return r;
}

function _kmFcd(ss) {
  var aba = getAba(ss, ABA_FCD);
  if (!aba) return {};
  var dados = aba.getRange(2,1,Math.max(aba.getLastRow()-1,1),
              aba.getLastColumn()).getValues();
  var mapa  = _mapaPrefixoPlaca(ss);
  var r     = {};
  dados.forEach(function(l) {
    var pref = l[1] ? _norm(l[1]) : null;
    var km   = _parseKm(l[3]);
    var pl   = l[7] ? _norm(l[7]) : (mapa[pref]||null);
    if (pl && km > 0) r[pl] = {km:km, ts:l[0]};
  });
  console.log("  ✓ " + Object.keys(r).length + " de " + ABA_FCD);
  return r;
}

function _kmAbast(ss) {
  var aba = getAba(ss, ABA_ABAST);
  if (!aba) return {};
  var ul = aba.getLastRow();
  if (ul < 2) return {};
  var dados = aba.getRange(2,1,ul-1,aba.getLastColumn()).getValues();
  var r = {};
  dados.forEach(function(l) {
    var pl = l[ABAST_COL_PLACA] ? _norm(l[ABAST_COL_PLACA]) : null;
    var km = _parseKm(l[ABAST_COL_KM]);
    if (!pl || km < 1000 || km > 999999) return;
    if (!r[pl] || km > r[pl].km) r[pl] = {km:km, ts:l[ABAST_COL_DATA]};
  });
  console.log("  ✓ " + Object.keys(r).length + " de " + ABA_ABAST);
  return r;
}

function _gravarKm(ss, reg) {
  var total = 0;
  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var dados = aba.getDataRange().getValues();
    var cnt   = 0;
    for (var i=1; i<dados.length; i++) {
      var plNorm = _norm(dados[i][COL_PLACA]);
      var kma    = _parseKm(dados[i][COL_KM_ATUAL]);
      var pref   = String(dados[i][COL_PREFIXO]||"").trim();
      if (!plNorm) continue;
      var entrada = reg[plNorm] || null;
      if (!entrada) continue;
      if (entrada.km > kma) {
        aba.getRange(i+1, COL_KM_ATUAL+1).setValue(entrada.km);
        aba.getRange(i+1, 4).setValue(entrada.ts);
        cnt++; total++;
        console.log("  ✅ "+nomeAba+" | "+pref+" | "+plNorm+": "
          +kma.toLocaleString("pt-BR")+" → "+entrada.km.toLocaleString("pt-BR"));
      }
    }
    console.log("  "+nomeAba+": "+cnt+" KM atualizado(s)");
  });
  return total;
}

function _mapaPrefixoPlaca(ss) {
  var m = {};
  [ABA_1SGB, ABA_2SGB].forEach(function(n) {
    var aba = getAba(ss, n);
    if (!aba) return;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var pref  = _norm(d[i][COL_PREFIXO]);
      var placa = _norm(d[i][COL_PLACA]);
      if (pref && placa) m[pref] = placa;
    }
  });
  return m;
}

// ============================================================
// PARTE 2: STATUS
// ============================================================
function sincronizarStatus(ss) {
  var prefixos = _extrairPrefixos17GB(ss);
  var sp = {};
  _combinarStatus(sp, _lerStatusFiltrado(ID_PLANILHA_OPERACIONAL, "US SEM OPERACAO", prefixos));
  _combinarStatus(sp, _lerStatusFiltrado(ID_PLANILHA_OPERACIONAL, "US OPERANDO",     prefixos));
  _combinarStatus(sp, _lerStatusFiltrado(ID_CBM,                  "US OPERANDO CBM", prefixos));
  if (!Object.keys(sp).length)
    return {atualizados:0, baixada:0, operando:0, reserva:0, divergencias:[]};
  return _gravarStatus(ss, sp);
}

function _extrairPrefixos17GB(ss) {
  var p = {};
  [ABA_1SGB, ABA_2SGB].forEach(function(n) {
    var aba = getAba(ss, n);
    if (!aba) return;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var pref = _norm(d[i][COL_PREFIXO]);
      if (pref) p[pref] = true;
    }
  });
  return Object.keys(p);
}

function _lerStatusFiltrado(idPlanilha, nomeAba, prefixosAlvo) {
  var m = {};
  try {
    var ss   = SpreadsheetApp.openById(idPlanilha);
    var abas = ss.getSheets();
    var aba  = null;
    for (var i=0; i<abas.length; i++) {
      if (abas[i].getName().toUpperCase().indexOf(nomeAba.toUpperCase()) >= 0) {
        aba = abas[i]; break;
      }
    }
    if (!aba) return m;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var p = _norm(d[i][0]);
      var s = String(d[i][6]||"").trim();
      if (prefixosAlvo.indexOf(p) >= 0 && s) {
        if (MAPA_STATUS[s]) s = MAPA_STATUS[s];
        m[p] = s;
      }
    }
  } catch(e) { console.error("  ❌ "+nomeAba+": "+e.message); }
  return m;
}

function _gravarStatus(ss, sp) {
  var tot=0, bx=0, op=0, rv=0, divergencias=[];

  function _cnt(st, delta) {
    var u = (st||"").toUpperCase();
    if      (u.indexOf("BAIXA")   >= 0) bx += delta;
    else if (u.indexOf("RESERVA") >= 0) rv += delta;
    else                                 op += delta;
  }

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var d   = aba.getDataRange().getValues();
    var cnt = 0;
    for (var i=1; i<d.length; i++) {
      var prefNorm   = _norm(d[i][COL_PREFIXO]);
      var prefOrig   = String(d[i][COL_PREFIXO]||"").trim().toUpperCase();
      var statusNovo = sp[prefNorm] || sp[prefOrig] || null;
      var statusAt   = String(d[i][7]||"").trim();
      if (!statusNovo) continue;

      if (statusAt && statusAt !== statusNovo) {
        divergencias.push({prefixo:prefOrig, aba:nomeAba,
                           statusAtual:statusAt, statusNovo:statusNovo});
        var statusFinal;
        if (_prioridade(statusNovo) >= _prioridade(statusAt)) {
          aba.getRange(i+1,8).setValue(statusNovo);
          aba.getRange(i+1,COL_P+1).setValue(statusNovo);
          statusFinal = statusNovo; cnt++; tot++;
        } else {
          if (String(d[i][COL_P]||"").trim() !== statusAt)
            aba.getRange(i+1,COL_P+1).setValue(statusAt);
          statusFinal = statusAt;
        }
        _cnt(statusFinal, 1);
        continue;
      }
      if (statusAt) _cnt(statusAt, 1);
      if (statusAt !== statusNovo) {
        if (statusAt) _cnt(statusAt,-1);
        _cnt(statusNovo, 1);
        aba.getRange(i+1,8).setValue(statusNovo);
        aba.getRange(i+1,COL_P+1).setValue(statusNovo);
        cnt++; tot++;
      } else {
        if (String(d[i][COL_P]||"").trim() !== statusNovo)
          aba.getRange(i+1,COL_P+1).setValue(statusNovo);
      }
    }
    console.log("  "+nomeAba+": "+cnt+" status atualizado(s)");
  });
  console.log("  📊 Baixadas:"+bx+" Operando:"+op+" Reserva:"+rv);
  return {atualizados:tot, baixada:bx, operando:op, reserva:rv, divergencias:divergencias};
}

// ============================================================
// PARTE 3: TAREFAS
// ============================================================
function lerTarefas(ss) {
  var r = {total:0, pendente:0, andamento:0, concluida:0, outros:0};
  var aba = getAba(ss, ABA_TAREFAS);
  if (!aba) return r;
  var ul = aba.getLastRow();
  if (ul < 2) return r;
  aba.getRange(2,1,ul-1,5).getValues().forEach(function(ln) {
    var desc = String(ln[2]||"").trim();
    if (!desc) return;
    var st = String(ln[4]||"").trim().toUpperCase()
               .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
    r.total++;
    if      (st==="PENDENTE")     r.pendente++;
    else if (st==="EM ANDAMENTO") r.andamento++;
    else if (st==="CONCLUIDA")    r.concluida++;
    else if (st)                  r.outros++;
  });
  console.log("  ✓ Tarefas:"+r.total+" Pendente:"+r.pendente
    +" Andamento:"+r.andamento+" Concluída:"+r.concluida);
  return r;
}

// ============================================================
// PARTE 4: ALERTAS L/M/N/O
// ============================================================
function calcularAlertas(ss) {
  var alertas = {
    bateria:   {vencidos:[], aVencer:[]},
    lavagem:   {vencidos:[], aVencer:[]},
    pneu:      {vencidos:[], aVencer:[]},
    embreagem: {vencidos:[], aVencer:[]},
    statusVencido:[], statusAVencer:[],
    total:0
  };
  var hoje = new Date(); hoje.setHours(0,0,0,0);

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var ul = aba.getLastRow();
    if (ul < 2) return;
    var numCols = Math.max(aba.getLastColumn(), 16);
    var dados   = aba.getRange(2,1,ul-1,numCols).getValues();

    dados.forEach(function(ln) {
      var pref = String(ln[COL_PREFIXO]||"").trim();
      var kmAt = _parseKm(ln[COL_KM_ATUAL]);
      if (!pref) return;

      // H — STATUS texto
      var valH = ln[COL_STATUS_H];
      if (valH && !(valH instanceof Date) && typeof valH !== "number") {
        var valHU = String(valH).trim().toUpperCase();
        if (valHU.indexOf("VENCIDO")  >= 0) { alertas.statusVencido.push({prefixo:pref,aba:nomeAba}); alertas.total++; }
        if (valHU.indexOf("A VENCER") >= 0) { alertas.statusAVencer.push({prefixo:pref,aba:nomeAba}); alertas.total++; }
      }

      // L — BATERIA (data ou texto)
      var valL = ln[COL_L];
      if (valL && valL !== "") {
        var dtBat = _parseData(valL);
        if (dtBat) {
          var db = Math.floor((dtBat - hoje)/(1000*60*60*24));
          if (db < 0)  { alertas.bateria.vencidos.push({prefixo:pref,aba:nomeAba,info:"Venceu "+_fmtData(dtBat)}); alertas.total++; }
          else if (db <= 30) { alertas.bateria.aVencer.push({prefixo:pref,aba:nomeAba,info:"Vence em "+db+"d"}); alertas.total++; }
        } else {
          var valLU = String(valL).trim().toUpperCase();
          if (valLU.indexOf("VENCIDO")  >= 0) { alertas.bateria.vencidos.push({prefixo:pref,aba:nomeAba,info:valLU}); alertas.total++; }
          if (valLU.indexOf("A VENCER") >= 0) { alertas.bateria.aVencer.push({prefixo:pref,aba:nomeAba,info:valLU}); alertas.total++; }
        }
      }

      // M — LAVAGEM (data)
      var valM = ln[COL_M];
      if (valM && valM !== "") {
        var dtLav = _parseData(valM);
        if (dtLav) {
          var dd = Math.floor((hoje - dtLav)/(1000*60*60*24));
          var df = LAVAGEM_DIAS - dd;
          if (df <= 0) { alertas.lavagem.vencidos.push({prefixo:pref,aba:nomeAba,ultimaData:_fmtData(dtLav),diasDesde:dd}); alertas.total++; }
          else if (df <= ALERTA_DIAS_AVISO) { alertas.lavagem.aVencer.push({prefixo:pref,aba:nomeAba,ultimaData:_fmtData(dtLav),diasFaltam:df}); alertas.total++; }
        }
      }

      // N — PNEU (km)
      var valN = ln[COL_N];
      if (valN && valN !== "" && kmAt > 0) {
        var kmP = _parseKm(valN);
        if (kmP > 0) {
          var fp = kmP - kmAt;
          if      (fp <= 0)             { alertas.pneu.vencidos.push({prefixo:pref,aba:nomeAba,kmFaltam:fp}); alertas.total++; }
          else if (fp <= ALERTA_KM_AVISO) { alertas.pneu.aVencer.push({prefixo:pref,aba:nomeAba,kmFaltam:fp}); alertas.total++; }
        }
      }

      // O — EMBREAGEM (km)
      var valO = ln[COL_O];
      if (valO && valO !== "" && kmAt > 0) {
        var kmE = _parseKm(valO);
        if (kmE > 0) {
          var fe = kmE - kmAt;
          if      (fe <= 0)             { alertas.embreagem.vencidos.push({prefixo:pref,aba:nomeAba,kmFaltam:fe}); alertas.total++; }
          else if (fe <= ALERTA_KM_AVISO) { alertas.embreagem.aVencer.push({prefixo:pref,aba:nomeAba,kmFaltam:fe}); alertas.total++; }
        }
      }
    });
  });

  console.log("  📊 Total alertas: "+alertas.total);
  return alertas;
}

function testarAlertas() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var tarefas      = lerTarefas(ss);
  var alertas      = calcularAlertas(ss);
  var alertasAbast = calcularAlertasAbastecimento(ss);
  var status       = {atualizados:0, baixada:0, operando:0, reserva:0, divergencias:[]};
  enviarTelegram(0, status, tarefas, alertas, alertasAbast);
}

// ============================================================
// PARTE 5: ABASTECIMENTO
// ============================================================


function _listarPlacasMonitoradas(ss) {
  var lista = [];
  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var d = aba.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      var pref  = String(d[i][COL_PREFIXO]||"").trim().toUpperCase();
      var placa = String(d[i][COL_PLACA]  ||"").trim().toUpperCase();
      if (pref && placa) lista.push({prefixo:pref, placa:placa, aba:nomeAba});
    }
  });
  return lista;
}
// Ignora datas futuras (> hoje + 30 dias) — evita datas erradas
// e usa CARIMBO (col A) como critério de "última VTR"
// pois o carimbo é automático e confiável
function calcularAlertasAbastecimento(ss) {
  var resultado = {
    totalDia:0, ultimaVtr:"", ultimaData:"",
    ultimaPlaca:"", atrasados:[], total:0
  };

  var aba = getAba(ss, ABA_ABAST);
  if (!aba) return resultado;
  var ul = aba.getLastRow();
  if (ul < 2) return resultado;

  var dados = aba.getRange(2,1,ul-1,aba.getLastColumn()).getValues();
  var hoje  = new Date(); hoje.setHours(0,0,0,0);
  var limite = new Date(hoje.getTime() + 30*24*60*60*1000); // hoje + 30d

  var mapaData       = {};
  var ultimaCarimbo  = null;
  var ultimaIdx      = -1;

  dados.forEach(function(l, idx) {
    var carimbo = l[ABAST_COL_CARIMBO]; // col A — Date automático confiável
    var dtReal  = _parseData(l[ABAST_COL_DATA]); // col H — data informada
    var pl      = l[ABAST_COL_PLACA]   ? _norm(l[ABAST_COL_PLACA])   : "";
    var pref    = l[ABAST_COL_PREFIXO] ? _norm(l[ABAST_COL_PREFIXO]) : "";

    if (!pl && !pref) return;

    // Usa carimbo para contar hoje e identificar última VTR
    if (carimbo instanceof Date) {
      var cs = new Date(carimbo); cs.setHours(0,0,0,0);

      // Conta hoje
      if (cs.getTime() === hoje.getTime()) resultado.totalDia++;

      // Última VTR = maior carimbo (automático, confiável)
      if (!ultimaCarimbo || carimbo > ultimaCarimbo) {
        ultimaCarimbo = carimbo;
        ultimaIdx     = idx;
      }
    }

    // Para o mapa de datas, usa col H SE válida e não futura
    if (dtReal && dtReal <= limite) {
      [pl, pref].forEach(function(chave) {
        if (!chave) return;
        if (!mapaData[chave] || dtReal > mapaData[chave]) mapaData[chave] = dtReal;
      });
    }
  });

  // Última VTR pelo carimbo mais recente
  if (ultimaIdx >= 0) {
    var ul2 = dados[ultimaIdx];
    // Data exibida = col H se válida, senão carimbo
    var dtExib = _parseData(ul2[ABAST_COL_DATA]);
    if (!dtExib || dtExib > limite) dtExib = _parseData(ul2[ABAST_COL_CARIMBO]);
    resultado.ultimaData  = _fmtData(dtExib);
    resultado.ultimaVtr   = ul2[ABAST_COL_PREFIXO]
      ? String(ul2[ABAST_COL_PREFIXO]).trim().toUpperCase() : "";
    resultado.ultimaPlaca = ul2[ABAST_COL_PLACA]
      ? String(ul2[ABAST_COL_PLACA]).trim().toUpperCase() : "";
  }

  console.log("  📊 Última VTR: "+resultado.ultimaVtr
    +" ("+resultado.ultimaPlaca+") — "+resultado.ultimaData);
  console.log("  📊 Abast. hoje: "+resultado.totalDia);

  // Verifica viaturas monitoradas
  _listarPlacasMonitoradas(ss).forEach(function(item) {
    var plNorm   = _norm(item.placa);
    var prefNorm = _norm(item.prefixo);
    var dtUlt    = mapaData[plNorm] || mapaData[prefNorm] || null;

    if (!dtUlt) {
      resultado.atrasados.push({
        prefixo:item.prefixo, placa:item.placa,
        ultimaData:"SEM REGISTRO", diasDesde:-1
      });
      resultado.total++;
      return;
    }
    var diasDesde = Math.floor((hoje - dtUlt)/(1000*60*60*24));
    if (diasDesde > ABAST_DIAS_ALERTA) {
      resultado.atrasados.push({
        prefixo:item.prefixo, placa:item.placa,
        ultimaData:_fmtData(dtUlt), diasDesde:diasDesde
      });
      resultado.total++;
    }
  });

  console.log("  📊 Alertas abast.: "+resultado.total);
  return resultado;
}

// _gravarKm — garante que ss é sempre passado corretamente
function _gravarKm(ss, reg) {
  if (!ss) { console.log("❌ _gravarKm: ss indefinido"); return 0; }
  var total = 0;

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var dados = aba.getDataRange().getValues();
    var cnt   = 0;

    for (var i=1; i<dados.length; i++) {
      var plNorm = _norm(dados[i][COL_PLACA]);
      var kma    = _parseKm(dados[i][COL_KM_ATUAL]);
      var pref   = String(dados[i][COL_PREFIXO]||"").trim();
      if (!plNorm) continue;

      var entrada = reg[plNorm] || null;
      if (!entrada) continue;

      if (entrada.km > kma) {
        aba.getRange(i+1, COL_KM_ATUAL+1).setValue(entrada.km);
        aba.getRange(i+1, 4).setValue(entrada.ts);
        cnt++; total++;
        console.log("  ✅ "+nomeAba+" | "+pref+" | "+plNorm
          +": "+kma.toLocaleString("pt-BR")
          +" → "+entrada.km.toLocaleString("pt-BR"));
      }
    }
    console.log("  "+nomeAba+": "+cnt+" KM atualizado(s)");
  });
  return total;
}
function testarAlertasAbastecimento() {
  var r = calcularAlertasAbastecimento(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(
    "⛽ ABASTECIMENTO\n\n"
    +"Última VTR: "+r.ultimaVtr+" ("+r.ultimaPlaca+") — "+r.ultimaData+"\n"
    +"Hoje: "+r.totalDia+" abastecimentos\n"
    +"Alertas: "+r.total
  );
}

// ============================================================
// TELEGRAM
// ============================================================
function enviarTelegram(totalKm, status, tarefas, alertas, alertasAbast) {
  var dt  = obterDataHoraBrasil();
  var msg = "📊 *FROTA 17ºGB — "+dt+"*\n\n";

  msg += "🔧 *KM ATUALIZADO:* "+totalKm+"\n";

  msg += SEP;
  msg += "📊 *STATUS OPERACIONAL*\n";
  msg += "🚒 Baixadas: "+status.baixada+"\n";
  msg += "🚗 Operando: "+status.operando+"\n";
  msg += "⏸️ Reserva:  "+status.reserva+"\n";

  if (tarefas.total > 0) {
    msg += SEP;
    msg += "📋 *TAREFAS*\n";
    msg += "📝 Total: "+tarefas.total+"\n";
    if (tarefas.pendente  > 0) msg += "🔴 PENDENTE: "     +tarefas.pendente+"\n";
    if (tarefas.andamento > 0) msg += "🟡 EM ANDAMENTO: " +tarefas.andamento+"\n";
    if (tarefas.concluida > 0) msg += "✅ CONCLUIDA: "    +tarefas.concluida+"\n";
    if (tarefas.outros    > 0) msg += "⚪ OUTROS: "       +tarefas.outros+"\n";
  }

  msg += SEP;
  if (alertas.total > 0) {
    msg += "⏰ *ALERTAS ("+alertas.total+")*\n";
    if (alertas.statusVencido.length   > 0) msg += "🚨 STATUS VENCIDO ("  +alertas.statusVencido.length+")\n";
    if (alertas.statusAVencer.length   > 0) msg += "⚠️ STATUS A VENCER (" +alertas.statusAVencer.length+")\n";
    if (alertas.bateria.vencidos.length   > 0) msg += "🚨 BATERIA ("   +alertas.bateria.vencidos.length+")\n";
    if (alertas.bateria.aVencer.length    > 0) msg += "⚠️ BATERIA ("   +alertas.bateria.aVencer.length+")\n";
    if (alertas.lavagem.vencidos.length   > 0) msg += "🚨 LAVAGEM ("   +alertas.lavagem.vencidos.length+")\n";
    if (alertas.lavagem.aVencer.length    > 0) msg += "⚠️ LAVAGEM ("   +alertas.lavagem.aVencer.length+")\n";
    if (alertas.pneu.vencidos.length      > 0) msg += "🚨 PNEU ("      +alertas.pneu.vencidos.length+")\n";
    if (alertas.pneu.aVencer.length       > 0) msg += "⚠️ PNEU ("      +alertas.pneu.aVencer.length+")\n";
    if (alertas.embreagem.vencidos.length > 0) msg += "🚨 EMBREAGEM (" +alertas.embreagem.vencidos.length+")\n";
    if (alertas.embreagem.aVencer.length  > 0) msg += "⚠️ EMBREAGEM (" +alertas.embreagem.aVencer.length+")\n";
  } else {
    msg += "✅ Sem alertas de manutencao!\n";
  }

  msg += SEP;
  msg += "⛽ *ABASTECIMENTO*\n";
  if (alertasAbast && alertasAbast.ultimaVtr) {
    // Apenas prefixo — sem placa
    msg += "🚒 Última VTR: *"+alertasAbast.ultimaVtr+"*"
      +" — "+alertasAbast.ultimaData+"\n";
  } else {
    msg += "🚒 Última VTR: sem registro\n";
  }
  msg += "📋 Abastecimentos hoje: *"+(alertasAbast ? alertasAbast.totalDia : 0)+"*\n";
  if (alertasAbast && alertasAbast.total > 0) {
    msg += "🚨 VTR sem abast. há +"+ABAST_DIAS_ALERTA+"d: *"+alertasAbast.total+"*\n";
  } else {
    msg += "✅ Todas as VTR abastecidas em dia!\n";
  }

  msg += SEP+"✅ *Fim do relatorio*";
  _postTelegram(msg);
}

function _postTelegram(texto) {
  var url    = "https://api.telegram.org/bot"+TOKEN_TELEGRAM+"/sendMessage";
  var LIMITE = 4000;
  var blocos = [];
  while (texto.length > 0) {
    if (texto.length <= LIMITE) { blocos.push(texto); break; }
    var corte = texto.lastIndexOf("\n", LIMITE);
    if (corte < 0) corte = LIMITE;
    blocos.push(texto.substring(0, corte));
    texto = texto.substring(corte).trim();
  }
  blocos.forEach(function(bloco, idx) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method:"post", contentType:"application/json",
        payload:JSON.stringify({chat_id:CHAT_ID, text:bloco, parse_mode:"Markdown"}),
        muteHttpExceptions:true
      });
      if (resp.getResponseCode() === 200) {
        console.log("📱 Telegram "+(idx+1)+"/"+blocos.length+" ✅");
      } else {
        UrlFetchApp.fetch(url, {
          method:"post", contentType:"application/json",
          payload:JSON.stringify({chat_id:CHAT_ID,
            text:bloco.replace(/\*/g,"").replace(/_/g,"")}),
          muteHttpExceptions:true
        });
        console.log("📱 Telegram "+(idx+1)+" (sem MD) ✅");
      }
      if (blocos.length > 1) Utilities.sleep(500);
    } catch(e) { console.log("❌ Telegram: "+e.message); }
  });
}

// ============================================================
// FIPE
// ============================================================
function atualizarTudo() {
  Logger.log("💰 FIPE + GASTOS...");
  buscarFIPE();
  atualizarGastos();
  Logger.log("✅ Concluído!");
}

function buscarFIPE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sf = getAba(ss, ABA_FROTA);
  var sg = getAba(ss, ABA_GASTOS);
  if (!sf || !sg) { Logger.log("⚠️ FROTA/GASTOS não encontrada"); return; }
  var dG = sg.getDataRange().getValues();
  var mapG = {};
  for (var i=1; i<dG.length; i++) {
    var p = dG[i][0] ? dG[i][0].toString().trim() : "";
    if (p) mapG[p] = i+1;
  }
  for (var row=2; row<=sf.getLastRow(); row++) {
    var pref = sf.getRange(row,1).getValue().toString().trim();
    var cod  = sf.getRange(row,2).getValue().toString().trim();
    var ano  = sf.getRange(row,11).getValue().toString().trim();
    if (!cod) continue;
    var val = _buscarFipe(cod, ano);
    if (val > 0) {
      var fmt = "R$"+val.toLocaleString("pt-BR",{minimumFractionDigits:2});
      if (sf.getRange(row,3).getValue().toString().trim() !== fmt) sf.getRange(row,3).setValue(fmt);
      if (mapG[pref] && sg.getRange(mapG[pref],2).getValue().toString().trim() !== fmt)
        sg.getRange(mapG[pref],2).setValue(fmt);
    }
    Utilities.sleep(800);
  }
}

function atualizarGastos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sg = getAba(ss, ABA_GASTOS);
  var sr = getAba(ss, ABA_RIV);
  if (!sg || !sr) { Logger.log("⚠️ GASTOS/RIV não encontrada"); return; }
  var anoC = new Date().getFullYear();
  var dR   = sr.getDataRange().getValues();
  var mapG = {};
  for (var i=1; i<dR.length; i++) {
    var pref = dR[i][RIV_COL_PREFIXO] ? dR[i][RIV_COL_PREFIXO].toString().trim() : "";
    var dv   = dR[i][RIV_COL_DATA];
    var vr   = dR[i][RIV_COL_VALOR]   ? dR[i][RIV_COL_VALOR].toString().trim()   : "";
    if (!pref) continue;
    var anoR = null;
    if (dv instanceof Date) anoR = dv.getFullYear();
    else if (typeof dv==="string" && dv.indexOf("/")>=0) anoR = parseInt(dv.split("/")[2],10);
    if (anoR !== anoC) continue;
    var vn = parseFloat((vr||"0").replace("R$","").replace(/\s/g,"").replace(/\./g,"").replace(",",".")) || 0;
    mapG[pref] = (mapG[pref]||0) + vn;
  }
  for (var row=2; row<=sg.getLastRow(); row++) {
    var p  = sg.getRange(row,1).getValue().toString().trim();
    if (!p) continue;
    var gt = mapG[p] || 0;
    sg.getRange(row,3).setValue("R$"+gt.toLocaleString("pt-BR",{minimumFractionDigits:2}));
    var fp = parseFloat(sg.getRange(row,2).getValue().toString()
      .replace("R$","").replace(/\s/g,"").replace(/\./g,"").replace(",",".")) || 0;
    sg.getRange(row,4).setValue(fp>0&&gt>0 ? ((gt/fp)*100).toFixed(2)+"%" : gt===0?"0,00%":"N/A");
    sg.getRange(row,5).setValue(fp>0?"✅ OK":"⚠️ Sem FIPE");
  }
  Logger.log("✅ GASTOS OK — Ano: "+anoC);
}

function _buscarFipe(cod, ano) {
  for (var ti=0; ti<TIPOS_API.length; ti++) {
    try {
      var ar = _fetchFipe(BASE_URL+"/"+TIPOS_API[ti]+"/"+cod+"/years");
      if (!ar||!ar.length) continue;
      var ao = null;
      for (var ai=0; ai<ar.length; ai++) {
        if (ar[ai].name.toString().indexOf(ano)>=0) { ao=ar[ai]; break; }
      }
      if (!ao) ao = ar[0];
      var vr = _fetchFipe(BASE_URL+"/"+TIPOS_API[ti]+"/"+cod+"/years/"+ao.code);
      if (!vr||!vr.price) continue;
      var pn = parseFloat(vr.price.toString().replace("R$","").replace(/\s/g,"").replace(/\./g,"").replace(",","."));
      if (!isNaN(pn)&&pn>0) { Utilities.sleep(600); return pn; }
    } catch(e) {}
  }
  return 0;
}

function _fetchFipe(url) {
  try {
    var r = UrlFetchApp.fetch(url, {
      method:"get",
      headers:{"X-Subscription-Token":TOKEN_FIPE},
      muteHttpExceptions:true
    });
    return r.getResponseCode()===200 ? JSON.parse(r.getContentText()) : null;
  } catch(e) { return null; }
}

// ============================================================
// DIAGNÓSTICOS
// ============================================================
function diagnosticarBateria() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var dados = aba.getRange(2,1,aba.getLastRow()-1,16).getValues();
    console.log("\n📋 "+nomeAba);
    dados.forEach(function(ln,i) {
      var pref = String(ln[COL_PREFIXO]||"").trim();
      var valL = ln[COL_L];
      if (!pref||!valL||valL==="") return;
      console.log("  L"+(i+2)+" | "+pref+" | L=["+valL+"] tipo="+typeof valL);
    });
  });
}

function diagnosticarMNO() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoje = new Date(); hoje.setHours(0,0,0,0);
  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var ul      = aba.getLastRow();
    var numCols = Math.max(aba.getLastColumn(),16);
    var dados   = aba.getRange(2,1,ul-1,numCols).getValues();
    console.log("\n📋 "+nomeAba);
    dados.forEach(function(ln,i) {
      var pref = String(ln[COL_PREFIXO]||"").trim();
      if (!pref) return;
      var vL=ln[COL_L], vM=ln[COL_M], vN=ln[COL_N], vO=ln[COL_O];
      if (!vL&&!vM&&!vN&&!vO) return;
      var kmAt = _parseKm(ln[COL_KM_ATUAL]);
      console.log("  L"+(i+2)+" | "+pref+" | km="+kmAt);
      if (vL&&vL!=="") { var dtL=_parseData(vL); console.log("    L=["+vL+"] → "+(dtL?_fmtData(dtL):String(vL).toUpperCase())); }
      if (vM&&vM!=="") { var dtM=_parseData(vM); if(dtM){var dd=Math.floor((hoje-dtM)/(864e5));console.log("    M=["+vM+"] há "+dd+"d "+(LAVAGEM_DIAS-dd<=0?"🚨":"✅"));}}
      if (vN&&vN!=="") { var kp=_parseKm(vN); console.log("    N=["+vN+"] faltam "+(kp-kmAt)+"km "+(kp-kmAt<=0?"🚨":"✅")); }
      if (vO&&vO!=="") { var ke=_parseKm(vO); console.log("    O=["+vO+"] faltam "+(ke-kmAt)+"km "+(ke-kmAt<=0?"🚨":"✅")); }
    });
  });
}

function limparColunasNO() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) return;
    var numCols = Math.max(aba.getLastColumn(),16);
    var dados   = aba.getRange(2,1,aba.getLastRow()-1,numCols).getValues();
    dados.forEach(function(ln,i) {
      var pref = String(ln[COL_PREFIXO]||"").trim();
      if (!pref) return;
      [COL_N,COL_O].forEach(function(col) {
        var v = ln[col];
        if (!v||v==="") return;
        var limpar = (typeof v==="string"&&(v.indexOf("|")>=0||v.indexOf("/")>=0))
                  || (v instanceof Date)
                  || (typeof v==="number"&&v>1000000);
        var kmValido = (typeof v==="number"&&v>=1000&&v<=999999);
        if (limpar&&!kmValido) {
          aba.getRange(i+2,col+1).setValue("");
          log.push(nomeAba+" L"+(i+2)+" "+pref+" col "+(col===COL_N?"N":"O")+": ["+v+"]");
        }
      });
    });
  });
  console.log("🧹 "+log.length+" células limpas");
  log.forEach(function(l){console.log("  "+l);});
}

// ============================================================
// DEBUG COMPLETO
// ============================================================
function debugCompleto() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var log  = [];

  log.push("╔══════════════════════════════════════════════╗");
  log.push("║  DEBUG 17ºGB v9.1 — "+new Date().toLocaleString("pt-BR"));
  log.push("╚══════════════════════════════════════════════╝");

  var abasEsperadas = [ABA_KM_VTR,ABA_FCD,ABA_ABAST,
                       ABA_1SGB,ABA_2SGB,ABA_FROTA,ABA_GASTOS,ABA_RIV,ABA_TAREFAS];
  log.push("\n📂 ABAS:");
  abasEsperadas.forEach(function(n) {
    log.push("  "+(getAba(ss,n)?"✅":"❌")+" \""+n+"\"");
  });

  var abaAbast = getAba(ss, ABA_ABAST);
  if (abaAbast) {
    var tl  = abaAbast.getLastRow()-1;
    var tc  = abaAbast.getLastColumn();
    var hd  = abaAbast.getRange(1,1,1,tc).getValues()[0];
    log.push("\n⛽ ABAST. VTR — "+tl+" linhas");
    log.push("  CARIMBO ["+_colLetra(ABAST_COL_CARIMBO)+"]: "+hd[ABAST_COL_CARIMBO]);
    log.push("  PREFIXO ["+_colLetra(ABAST_COL_PREFIXO)+"]: "+hd[ABAST_COL_PREFIXO]);
    log.push("  PLACA   ["+_colLetra(ABAST_COL_PLACA)  +"]: "+hd[ABAST_COL_PLACA]);
    log.push("  DATA    ["+_colLetra(ABAST_COL_DATA)   +"]: "+hd[ABAST_COL_DATA]);
    log.push("  KM      ["+_colLetra(ABAST_COL_KM)     +"]: "+hd[ABAST_COL_KM]);

    var dT = abaAbast.getRange(2,1,tl,tc).getValues();
    var v5=0,s5=0,sp5=0,sd5=0,sk5=0,th=0;
    var uDate=null,uVtr="",uPlaca="";
    dT.forEach(function(l) {
      var dt  = _parseData(l[ABAST_COL_DATA]);
      var pl  = l[ABAST_COL_PLACA]   ? String(l[ABAST_COL_PLACA]).trim()   : "";
      var pr  = l[ABAST_COL_PREFIXO] ? String(l[ABAST_COL_PREFIXO]).trim() : "";
      var km  = _parseKm(l[ABAST_COL_KM]);
      var car = l[ABAST_COL_CARIMBO];
      if (!pl)  sp5++;
      if (!pr)  
            if (!dt)  sd5++;
      if (!km)  sk5++;
      if (dt && (pl||pr)) {
        v5++;
        if (car instanceof Date) {
          var cs = new Date(car); cs.setHours(0,0,0,0);
          if (cs.getTime()===hoje.getTime()) th++;
        }
        if (!uDate||dt>uDate) { uDate=dt; uVtr=pr; uPlaca=pl; }
      }
    });
    log.push("  ✅ Válidos: "+v5+" | ❌ Sem placa: "+sp5
      +" | Sem data: "+sd5+" | Sem KM: "+sk5);
    log.push("  📅 Abast. hoje: "+th);
    log.push("  🚒 Última VTR: "+uVtr+" ("+uPlaca+") — "+_fmtData(uDate));

    log.push("  Últimas 5 linhas com dados:");
    var c5=0;
    for (var i=dT.length-1; i>=0&&c5<5; i--) {
      var pr2 = dT[i][ABAST_COL_PREFIXO] ? String(dT[i][ABAST_COL_PREFIXO]).trim() : "";
      var pl2 = dT[i][ABAST_COL_PLACA]   ? String(dT[i][ABAST_COL_PLACA]).trim()   : "";
      var dt2 = _parseData(dT[i][ABAST_COL_DATA]);
      var km2 = _parseKm(dT[i][ABAST_COL_KM]);
      if (!pr2&&!pl2) continue;
      log.push("    L"+(i+2)+" | "+pr2+" | "+pl2+" | "+_fmtData(dt2)+" | km="+km2);
      c5++;
    }
  }

  [ABA_1SGB, ABA_2SGB].forEach(function(nomeAba) {
    var aba = getAba(ss, nomeAba);
    if (!aba) { log.push("\n🚒 "+nomeAba+": ❌ não encontrada"); return; }
    var ul      = aba.getLastRow()-1;
    var numCols = Math.max(aba.getLastColumn(),16);
    var dados   = aba.getRange(2,1,Math.min(ul,999),numCols).getValues();
    var cP=0,cPl=0,cK=0,cH=0,cL=0,cM=0,cN=0,cO=0;
    dados.forEach(function(ln) {
      if (String(ln[COL_PREFIXO] ||"").trim()) cP++;
      if (String(ln[COL_PLACA]   ||"").trim()) cPl++;
      if (_parseKm(ln[COL_KM_ATUAL]))          cK++;
      if (String(ln[COL_STATUS_H]||"").trim()) cH++;
      if (ln[COL_L]&&ln[COL_L]!=="")          cL++;
      if (ln[COL_M]&&ln[COL_M]!=="")          cM++;
      if (ln[COL_N]&&ln[COL_N]!=="")          cN++;
      if (ln[COL_O]&&ln[COL_O]!=="")          cO++;
    });
    log.push("\n🚒 "+nomeAba+" — "+ul+" linhas");
    log.push("  A:"+cP+" B:"+cPl+" C:"+cK+" H:"+cH
      +" L:"+cL+" M:"+cM+" N:"+cN+" O:"+cO);
    var c3=0;
    log.push("  Amostra (3 primeiras):");
    for (var i=0; i<dados.length&&c3<3; i++) {
      var p=String(dados[i][COL_PREFIXO]||"").trim();
      if (!p) continue; c3++;
      log.push("    L"+(i+2)+": "+p
        +" | "+String(dados[i][COL_PLACA]||"").trim()
        +" | km="+_parseKm(dados[i][COL_KM_ATUAL])
        +" | H=["+String(dados[i][COL_STATUS_H]||"").trim()+"]"
        +" | L=["+dados[i][COL_L]+"]"
        +" | M=["+dados[i][COL_M]+"]"
        +" | N=["+dados[i][COL_N]+"]"
        +" | O=["+dados[i][COL_O]+"]");
    }
  });

  var tar = lerTarefas(ss);
  log.push("\n📋 TAREFAS: Total:"+tar.total+" Pendente:"+tar.pendente
    +" Andamento:"+tar.andamento+" Concluída:"+tar.concluida);

  log.push("\n⚙️ CONFIGURAÇÕES:");
  log.push("  ABAST_COL_CARIMBO:"+ABAST_COL_CARIMBO+" ("+_colLetra(ABAST_COL_CARIMBO)+")");
  log.push("  ABAST_COL_PREFIXO:"+ABAST_COL_PREFIXO+" ("+_colLetra(ABAST_COL_PREFIXO)+")");
  log.push("  ABAST_COL_PLACA:  "+ABAST_COL_PLACA  +" ("+_colLetra(ABAST_COL_PLACA)+")");
  log.push("  ABAST_COL_DATA:   "+ABAST_COL_DATA   +" ("+_colLetra(ABAST_COL_DATA)+")");
  log.push("  ABAST_COL_KM:     "+ABAST_COL_KM     +" ("+_colLetra(ABAST_COL_KM)+")");
  log.push("  LAVAGEM_DIAS:     "+LAVAGEM_DIAS);
  log.push("  ABAST_DIAS_ALERTA:"+ABAST_DIAS_ALERTA);

  log.push("\n╔══════════════════════════════════════════════╗");
  log.push("║  ✅ DEBUG CONCLUÍDO");
  log.push("╚══════════════════════════════════════════════╝");

  log.forEach(function(l){ console.log(l); });

  var msg = "🐛 *DEBUG 17ºGB v9.1*\n"+SEP;
  msg += "📂 *ABAS:*\n";
  abasEsperadas.forEach(function(n) {
    msg += (getAba(ss,n)?"✅":"❌")+" "+n+"\n";
  });
  if (abaAbast) {
    var hd2 = abaAbast.getRange(1,1,1,abaAbast.getLastColumn()).getValues()[0];
    msg += SEP+"⛽ *ABAST. VTR*\n";
    msg += "Linhas: "+(abaAbast.getLastRow()-1)+"\n";
    msg += "CARIMBO ["+_colLetra(ABAST_COL_CARIMBO)+"]: "+hd2[ABAST_COL_CARIMBO]+"\n";
    msg += "PREFIXO ["+_colLetra(ABAST_COL_PREFIXO)+"]: "+hd2[ABAST_COL_PREFIXO]+"\n";
    msg += "PLACA   ["+_colLetra(ABAST_COL_PLACA)  +"]: "+hd2[ABAST_COL_PLACA]+"\n";
    msg += "DATA    ["+_colLetra(ABAST_COL_DATA)   +"]: "+hd2[ABAST_COL_DATA]+"\n";
    msg += "KM      ["+_colLetra(ABAST_COL_KM)     +"]: "+hd2[ABAST_COL_KM]+"\n";
  }
  msg += SEP+"_Ver log: Apps Script → Execuções_";
  _postTelegram(msg);

  SpreadsheetApp.getUi().alert(
    "✅ DEBUG CONCLUÍDO\n\nVeja em:\nApps Script → Execuções → Ver logs\n\nResumo enviado ao Telegram."
  );
}

// ============================================================
// AGENDAMENTO
// ============================================================
function agendar5min() {
  removerTriggers();
  ScriptApp.newTrigger("sincronizarFrota")
    .timeBased().everyMinutes(5).create();
  console.log("✅ Agendado a cada 5 minutos");
}

function removerTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (["sincronizarFrota","atualizarTudo"]
        .indexOf(t.getHandlerFunction()) >= 0)
      ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// ATALHOS
// ============================================================
function testar()     { sincronizarFrota(); }
function testarFIPE() { atualizarTudo(); }
// ============================================================
// CACHE DE VIATURAS — Bot Telegram
// ============================================================
var _cacheVtrs = {};

function _listarVtrsAbaCached(nomeAba) {
  // Usa cache em memória durante a mesma execução
  if (_cacheVtrs[nomeAba]) {
    console.log("✅ Cache hit: "+nomeAba);
    return _cacheVtrs[nomeAba];
  }
  var lista = _listarVtrsAba(nomeAba);
  _cacheVtrs[nomeAba] = lista;
  return lista;
}

function _limparCacheVtrs() {
  _cacheVtrs = {};
  console.log("🧹 Cache viaturas limpo");
}