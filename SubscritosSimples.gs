// =============================================================
//  IMPORTAÇÃO SUBSCRITOS_8774  (append-only / incremental)
//
//  Origem : COMERCIAL da planilha 1XcHlMYLTHkye3Pp_qn6b9r9jqu7BwOztzVOUG1wfjOU
//  Destino: aba SUBSCRITOS_8774 da planilha 146o1jU6VRO35XbTzWVY587hP0vy2s1SK35m5TT33H7U
//
//  Regra: importa as colunas F, G, H, L, AU, BC, BV
//         se, e somente se,  B = 8774  E  C ∈ {MATRICULADO, UPGRADE}.
//  Cabeçalho da origem na linha 3 (dados a partir da linha 4).
//  A aba de destino é criada na 1ª importação, se não existir.
//  "Atualizar" = acrescenta apenas linhas novas, preservando as antigas.
// =============================================================

const SUB8774 = {
  ID_ORIGEM:   '1XcHlMYLTHkye3Pp_qn6b9r9jqu7BwOztzVOUG1wfjOU',
  ID_DESTINO:  '146o1jU6VRO35XbTzWVY587hP0vy2s1SK35m5TT33H7U',

  ABA_ORIGEM:  'COMERCIAL',
  ABA_DESTINO: 'SUBSCRITOS_8774',

  LINHA_CABECALHO: 3,   // cabeçalho da origem
  LINHA_DADOS:     4,   // 1ª linha de dados da origem

  // Colunas da origem a exportar (0-based):
  // F=5, G=6, H=7, L=11, AU=46, BC=54, BV=73
  INDICES_COLUNAS: [5, 6, 7, 11, 46, 54, 73],
  MAX_COLUNAS:     74,  // largura de leitura (precisa alcançar BV = 74)

  // Filtro (0-based):
  FILTRO_B_IDX:  1,                            // coluna B
  FILTRO_B_VAL:  '8774',                       // B = 8774
  FILTRO_C_IDX:  2,                            // coluna C
  FILTRO_C_OK:   ['MATRICULADO', 'UPGRADE'],   // C ∈ {MATRICULADO, UPGRADE}
};

const SUB8774_CHUNK = 1000;   // linhas lidas por vez da origem

// =============================================================
//  MENU
// =============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('➡️ Importações')
    .addItem('Atualizar SUBSCRITOS_8774 (apenas novos)', 'atualizarSubscritos8774')
    .addToUi();
}

// =============================================================
//  FUNÇÃO PRINCIPAL — atualização incremental (append-only)
// =============================================================

function atualizarSubscritos8774() {
  const NOME = 'atualizarSubscritos8774';
  const cfg  = SUB8774;

  // Evita execuções concorrentes (menu + eventual gatilho).
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    Logger.log('[' + NOME + '] Lock ocupado. Abortado.');
    return;
  }

  try {
    const abaOrigem  = _sub8774_abrirAba(cfg.ID_ORIGEM, cfg.ABA_ORIGEM);
    const abaDestino = _sub8774_abrirOuCriarAba(cfg.ID_DESTINO, cfg.ABA_DESTINO);

    const ultimaLinha = abaOrigem.getLastRow();
    const larguraLer  = Math.min(cfg.MAX_COLUNAS, abaOrigem.getLastColumn());

    // Cabeçalho no destino (só se a aba estiver vazia).
    if (abaDestino.getLastRow() === 0) {
      const cab      = abaOrigem.getRange(cfg.LINHA_CABECALHO, 1, 1, larguraLer).getValues()[0];
      const cabFinal = cfg.INDICES_COLUNAS.map(function (i) { return i < cab.length ? cab[i] : ''; });
      abaDestino.getRange(1, 1, 1, cabFinal.length).setValues([cabFinal]);
      SpreadsheetApp.flush();
    }

    if (ultimaLinha < cfg.LINHA_DADOS) {
      _sub8774_avisar('ℹ️ A origem (COMERCIAL) não tem dados a partir da linha ' + cfg.LINHA_DADOS + '.');
      return;
    }

    // Chaves já existentes no destino (para não duplicar).
    const existentes = _sub8774_lerChavesDestino(abaDestino, cfg.INDICES_COLUNAS.length);

    // Lê a origem em blocos, filtra e acumula apenas o que ainda não existe.
    const novos    = [];
    const vistos    = {};   // dedup dentro desta própria execução
    let linhaAtual = cfg.LINHA_DADOS;
    let casaram    = 0;

    while (linhaAtual <= ultimaLinha) {
      const linhasNoChunk = Math.min(SUB8774_CHUNK, ultimaLinha - linhaAtual + 1);
      const chunk = abaOrigem.getRange(linhaAtual, 1, linhasNoChunk, larguraLer).getValues();

      chunk.forEach(function (l) {
        const b = String(l[cfg.FILTRO_B_IDX]).trim();
        const c = String(l[cfg.FILTRO_C_IDX]).trim().toUpperCase();
        if (b !== cfg.FILTRO_B_VAL) return;
        if (cfg.FILTRO_C_OK.indexOf(c) === -1) return;
        casaram++;

        const saida = cfg.INDICES_COLUNAS.map(function (i) { return i < l.length ? l[i] : ''; });
        const chave = _sub8774_chave(saida);
        if (existentes[chave] || vistos[chave]) return;   // já existe → não reimporta
        vistos[chave] = true;
        novos.push(saida);
      });

      linhaAtual += SUB8774_CHUNK;
    }

    if (novos.length > 0) {
      const proxLinha = abaDestino.getLastRow() + 1;
      abaDestino.getRange(proxLinha, 1, novos.length, cfg.INDICES_COLUNAS.length).setValues(novos);
      SpreadsheetApp.flush();
    }

    _sub8774_avisar('✅ SUBSCRITOS_8774 atualizado!\n' +
      novos.length + ' novo(s) registro(s) adicionado(s).\n' +
      casaram + ' linha(s) casaram com o filtro (B = 8774 e C ∈ {MATRICULADO, UPGRADE}); ' +
      'as já existentes foram mantidas.');

  } catch (e) {
    Logger.log('[ERRO] ' + NOME + ': ' + e.message + '\n' + e.stack);
    _sub8774_avisar('❌ Erro: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
//  AUXILIARES
// =============================================================

/** Lê as chaves (linha inteira normalizada) já presentes no destino. */
function _sub8774_lerChavesDestino(abaDestino, largura) {
  const chaves = {};
  const ultima = abaDestino.getLastRow();
  if (ultima < 2) return chaves;   // linha 1 = cabeçalho
  const dados = abaDestino.getRange(2, 1, ultima - 1, largura).getValues();
  dados.forEach(function (r) { chaves[_sub8774_chave(r)] = true; });
  return chaves;
}

/** Chave de identidade de uma linha: junção normalizada de todos os valores. */
function _sub8774_chave(valores) {
  return valores
    .map(function (v) { return String(v == null ? '' : v).trim().toUpperCase(); })
    .join('||');
}

/** Abre uma aba de uma planilha por ID (erro claro se falhar). */
function _sub8774_abrirAba(idPlanilha, nomeAba) {
  let planilha;
  try { planilha = SpreadsheetApp.openById(idPlanilha); }
  catch (e) {
    throw new Error('Não foi possível abrir a planilha ' + idPlanilha +
      '. Verifique o ID e a permissão de acesso.');
  }
  const aba = planilha.getSheetByName(nomeAba);
  if (!aba) throw new Error("Aba '" + nomeAba + "' não encontrada em " + idPlanilha + '.');
  return aba;
}

/** Abre a aba de destino por ID; cria-a se ainda não existir. */
function _sub8774_abrirOuCriarAba(idPlanilha, nomeAba) {
  let planilha;
  try { planilha = SpreadsheetApp.openById(idPlanilha); }
  catch (e) {
    throw new Error('Não foi possível abrir a planilha de destino ' + idPlanilha +
      '. Verifique o ID e a permissão de acesso.');
  }
  let aba = planilha.getSheetByName(nomeAba);
  if (!aba) aba = planilha.insertSheet(nomeAba);
  return aba;
}

/** Alerta na UI quando há contexto (menu); silencioso em gatilhos. */
function _sub8774_avisar(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); }
  catch (e) { Logger.log(mensagem); }
}
