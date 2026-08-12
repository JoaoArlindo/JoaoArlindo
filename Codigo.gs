// =============================================================
//  CONFIGURAÇÃO GLOBAL
// =============================================================

const CONFIG = {
  ID_ORIGEM: '1XcHlMYLTHkye3Pp_qn6b9r9jqu7BwOztzVOUG1wfjOU',

  SUBSCRITOS: {
    ABA_ORIGEM:        'COMERCIAL',
    ABA_DESTINO:       'SUBSCRITOS_2025',
    LINHA_CABECALHO:   3,
    LINHA_DADOS:       4,
    MAX_COLUNAS:       89,
    // AU (46 = NUM WATSHAPP) e BV (73 = Perfil no Instagram?) removidas do import.
    INDICES_COLUNAS:   [1, 2, 5, 6, 7, 11, 12, 85, 83, 88],
    FILTRO_STATUS_IDX: 2,
    FILTRO_TURMA_IDX:  1,
    FILTRO_STATUS_OK:  ['MATRICULADO', 'UPGRADE'],
    FILTRO_TURMA_OK:   '7388',
  },

  // ---- IMPORTAÇÃO 2020: origem é OUTRA planilha (não a JETHRO COMERCIAL) ----
  SUBSCRITOS_2020: {
    ID_ORIGEM:         '1KjTMgpe1m9TftwCd2Wtwfr59z2fyGiwxHIDZCHLtokU',
    ABA_ORIGEM:        'COMERCIAL',
    ABA_DESTINO:       'SUBSCRITOS_2020',
    LINHA_CABECALHO:   4,
    LINHA_DADOS:       5,
    MAX_COLUNAS:       79,    // precisa alcançar a coluna CA (=79) pedida na saída
    INDICES_COLUNAS:   [1, 2, 5, 6, 7, 0, 44, 78],   // saída: B, C, F, G, H, A, AS, CA
    FILTRO_COL_IDX:    1,     // coluna B (COD VDDOR), 0-based
    FILTRO_TXT:        '7388',// B = 7388 (token; aceita "7388" e "7388 / 8357")
    FILTRO_STATUS_IDX: 2,     // coluna C (status), 0-based
    FILTRO_STATUS_OK:  ['MATRICULADO', 'UPGRADE'],
  },

  DIRETORES: {
    ABA_ORIGEM:      'DIRECTORS',
    ABA_DESTINO:     'DIRETORES',
    SRC_LINHA_DADOS: 2,
    DST_LINHA_DADOS: 2,
    SRC_NUM_COLS:    13,     // largura FIXA de propósito (evita timeout)
    DST_NUM_COLS:    6,      // DIRETORES = A(ID)..F(DATA CONDECORAÇÃO) — REDE SOCIAL (G) removida
    FILTRO_COL_IDX:  6,      // col G (DIRETOR) na origem
    FILTRO_TXT:      'JANINE BARBOSA - ID 7388',  // frase específica: sem falso positivo numérico

    CHAVE_ID_SRC_IDX:    2,
    CHAVE_NOME_SRC_IDX:  3,
    CHAVE_NOMEN_SRC_IDX: 4,

    // Índices no DESTINO (0-based) usados pela dedução "manter a melhor" (D2)
    DST_COND_IDX:    4,      // E FOI CONDECORADO?
    DST_DATA_IDX:    5,      // F DATA CONDECORAÇÃO

    MAP: [
      { src: 2,  dst: 1 },   // C ID                → A ID
      { src: 3,  dst: 2 },   // D NOME              → B NOME
      { src: 4,  dst: 3 },   // E NOMENCLATURA      → C NOMENCLATURA
      { src: 6,  dst: 4 },   // G DIRETOR           → D DIRETOR
      { src: 9,  dst: 5 },   // J FOI CONDECORADO?  → E FOI CONDECORADO?
      { src: 10, dst: 6 },   // K DATA CONDECORAÇÃO → F DATA CONDECORAÇÃO
      // REDE SOCIAL (G) removida — a aba DIRETORES vai apenas até F.
    ],
  },
};

const CHUNK_SIZE    = 500;
const MAX_EXEC_MS   = 270000;        // 4,5 min (margem antes do limite de 6 min)
const ESTADO_TTL_MS = 15 * 60 * 1000; // R2: idade máxima de um estado de continuação

// =============================================================
//  MENU
// =============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('➡️ Importações')
    .addItem('Atualizar Subscritos 2025',         'importarSubscritos')
    .addItem('Importar Subscritos 2020',          'importarSubscritos2020')
    .addItem('Importar Novos Diretores',          'importarDiretores')
    .addItem('🧹 Remover Diretores Duplicados',    'limparDuplicadosDiretores')
    .addSeparator()
    .addItem('⏰ Ativar atualização automática',    'criarGatilhos')
    .addItem('🗑️ Remover atualização automática',   'removerGatilhos')
    .addToUi();
}

// =============================================================
//  GATILHOS
// =============================================================

function criarGatilhos() {
  removerGatilhos();
  ScriptApp.newTrigger('importarSubscritos').timeBased().atHour(1).everyDays(1).create();
  ScriptApp.newTrigger('importarDiretores').timeBased().atHour(2).everyDays(1).create();
  _avisar('✅ Gatilhos criados!\n\n• Subscritos: todo dia à 1h\n• Diretores: todo dia às 2h');
}

function removerGatilhos() {
  // R1: inclui 'importarSubscritos2020' para também apagar gatilhos de continuação dele.
  const funcoes = ['importarSubscritos', 'importarDiretores', 'importarSubscritos2020'];
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (funcoes.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  if (removidos > 0) _avisar('🗑️ ' + removidos + ' gatilho(s) removido(s).');
}

// =============================================================
//  IMPORTAÇÃO 1 — SUBSCRITOS_2025
//  R3: lock de execução (evita concorrência gatilho+continuação / manual+automático).
//  R2: estado de continuação com carimbo de tempo e TTL (descarta órfãos).
// =============================================================

function importarSubscritos() {
  const NOME      = 'importarSubscritos';
  const CHAVE_EST = 'SUBSCRITOS';

  // R3: adquire o lock. Se já há execução, reagenda a continuação (se pendente) e sai.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    const temEstado = _carregarEstado(CHAVE_EST).linhaAtual !== undefined;
    if (temEstado) _agendarContinuacao(NOME);
    Logger.log('[' + NOME + '] Lock ocupado. ' + (temEstado ? 'Continuação reagendada.' : 'Abortado.'));
    return;
  }

  const cfg       = CONFIG.SUBSCRITOS;
  const startTime = Date.now();
  try {
    let estado        = _carregarEstado(CHAVE_EST);
    let ehContinuacao = estado.linhaAtual !== undefined;

    // R2: estado antigo/órfão → recomeça do zero em vez de retomar dados obsoletos.
    if (ehContinuacao && (!estado.ts || (Date.now() - estado.ts > ESTADO_TTL_MS))) {
      Logger.log('[' + NOME + '] Estado órfão descartado. Reinício do zero.');
      _limparEstado(CHAVE_EST); estado = {}; ehContinuacao = false;
    }

    const abaOrigem   = _abrirAba(CONFIG.ID_ORIGEM, cfg.ABA_ORIGEM);
    const ultimaLinha = abaOrigem.getLastRow();
    const abaDestino  = _abrirAbaLocal(cfg.ABA_DESTINO);

    if (ultimaLinha < cfg.LINHA_DADOS) {
      _avisar('ℹ️ Não há dados para importar em COMERCIAL.');
      _limparEstado(CHAVE_EST);
      return;
    }

    if (!ehContinuacao) {
      const cabecalho = abaOrigem.getRange(cfg.LINHA_CABECALHO, 1, 1, cfg.MAX_COLUNAS).getValues()[0];
      const cabecalhoFinal = cfg.INDICES_COLUNAS.map(i => cabecalho[i]);
      abaDestino.getDataRange().clearContent();
      abaDestino.getRange(1, 1, 1, cabecalhoFinal.length).setValues([cabecalhoFinal]);
    }

    let linhaAtual     = ehContinuacao ? estado.linhaAtual     : cfg.LINHA_DADOS;
    let totalImportado = ehContinuacao ? estado.totalImportado : 0;

    while (linhaAtual <= ultimaLinha) {
      if (Date.now() - startTime > MAX_EXEC_MS) {
        _salvarEstado(CHAVE_EST, { linhaAtual: linhaAtual, totalImportado: totalImportado, ts: Date.now() });
        _agendarContinuacao(NOME);
        Logger.log('[' + NOME + '] Pausado na linha ' + linhaAtual + '. Continuação agendada.');
        return;
      }

      const linhasNoChunk = Math.min(CHUNK_SIZE, ultimaLinha - linhaAtual + 1);
      const chunk = abaOrigem.getRange(linhaAtual, 1, linhasNoChunk, cfg.MAX_COLUNAS).getValues();

      const dadosFiltrados = chunk
        .filter(function (l) {
          return cfg.FILTRO_STATUS_OK.includes(l[cfg.FILTRO_STATUS_IDX]) &&
                 String(l[cfg.FILTRO_TURMA_IDX]) === cfg.FILTRO_TURMA_OK;
        })
        .map(function (l) { return cfg.INDICES_COLUNAS.map(function (i) { return l[i]; }); });

      if (dadosFiltrados.length > 0) {
        const proxLinha = abaDestino.getLastRow() + 1;
        abaDestino.getRange(proxLinha, 1, dadosFiltrados.length, dadosFiltrados[0].length).setValues(dadosFiltrados);
        totalImportado += dadosFiltrados.length;
      }

      linhaAtual += CHUNK_SIZE;
    }

    SpreadsheetApp.flush();
    _limparEstado(CHAVE_EST);
    _avisar('✅ Subscritos atualizados!\n' + totalImportado + ' registros importados.');

  } catch (e) {
    _logErro(NOME, e);
    _avisar('❌ Erro: ' + e.message);
  } finally {
    lock.releaseLock();   // R3: liberta o lock (inclusive nas saídas de continuação)
  }
}

// =============================================================
//  IMPORTAÇÃO 1B — SUBSCRITOS_2020  (cabeçalho na LINHA 4 da origem)
//  Escreve o cabeçalho + as linhas em que B = 7388 (token) E C ∈ {MATRICULADO, UPGRADE},
//  a partir de A1. Saída: B, C, F, G, H, A, AS, CA.
//  R3: lock de execução.  D3: filtro por token (_contemToken).
//  Diagnóstico anti-"some-e-volta": clearContents + deteção de fórmulas.
// =============================================================

function importarSubscritos2020() {
  const NOME      = 'importarSubscritos2020';
  const CHAVE_EST = 'SUBSCRITOS_2020';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    const temEstado = _carregarEstado(CHAVE_EST).linhaAtual !== undefined;
    if (temEstado) _agendarContinuacao(NOME);
    Logger.log('[' + NOME + '] Lock ocupado. ' + (temEstado ? 'Continuação reagendada.' : 'Abortado.'));
    return;
  }

  const cfg       = CONFIG.SUBSCRITOS_2020;
  const startTime = Date.now();
  try {
    let estado       = _carregarEstado(CHAVE_EST);
    let ehContinuacao = estado.linhaAtual !== undefined;
    if (ehContinuacao && (!estado.ts || (Date.now() - estado.ts > ESTADO_TTL_MS))) {
      Logger.log('[' + NOME + '] Estado órfão descartado. Reinício do zero.');
      _limparEstado(CHAVE_EST); estado = {}; ehContinuacao = false;
    }

    const abaOrigem  = _abrirAba(cfg.ID_ORIGEM, cfg.ABA_ORIGEM);
    const abaDestino = _abrirOuCriarAbaLocal(cfg.ABA_DESTINO);

    const ultimaLinha = abaOrigem.getLastRow();
    if (ultimaLinha < cfg.LINHA_CABECALHO) {
      _avisar('ℹ️ A origem 2020 não tem sequer a linha ' + cfg.LINHA_CABECALHO + ' (cabeçalho).');
      _limparEstado(CHAVE_EST);
      return;
    }

    const larguraLer = Math.min(cfg.MAX_COLUNAS, abaOrigem.getMaxColumns());

    let formulasAntes = [];
    if (!ehContinuacao) {
      formulasAntes = _detectarFormulas(abaDestino);   // causa provável do "revert"
      abaDestino.clearContents();                      // limpa VALORES E FÓRMULAS (preserva formatação)

      const cabecalho = abaOrigem.getRange(cfg.LINHA_CABECALHO, 1, 1, larguraLer).getValues()[0];
      const cabecalhoFinal = cfg.INDICES_COLUNAS.map(function (i) {
        return i < cabecalho.length ? cabecalho[i] : '';
      });
      abaDestino.getRange(1, 1, 1, cabecalhoFinal.length).setValues([cabecalhoFinal]);
      SpreadsheetApp.flush();   // getLastRow() confiável no loop
    }

    if (ultimaLinha < cfg.LINHA_DADOS) {
      SpreadsheetApp.flush();
      _limparEstado(CHAVE_EST);
      _limparTriggersDe(NOME);
      let m = '✅ Subscritos 2020: cabeçalho (linha ' + cfg.LINHA_CABECALHO + ') importado; sem dados.';
      if (formulasAntes.length) m += '\n\n⚠️ A aba continha fórmula(s): ' + formulasAntes.slice(0, 5).join(' | ');
      _avisar(m);
      return;
    }

    let linhaAtual     = ehContinuacao ? estado.linhaAtual     : cfg.LINHA_DADOS;
    let totalImportado = ehContinuacao ? estado.totalImportado : 0;

    while (linhaAtual <= ultimaLinha) {
      if (Date.now() - startTime > MAX_EXEC_MS) {
        _salvarEstado(CHAVE_EST, { linhaAtual: linhaAtual, totalImportado: totalImportado, ts: Date.now() });
        _agendarContinuacao(NOME);
        Logger.log('[' + NOME + '] Pausado na linha ' + linhaAtual + '. Continuação agendada.');
        return;
      }

      const linhasNoChunk = Math.min(CHUNK_SIZE, ultimaLinha - linhaAtual + 1);
      const chunk = abaOrigem.getRange(linhaAtual, 1, linhasNoChunk, larguraLer).getValues();

      // Filtro: B = 7388 (token; aceita "7388" e "7388 / 8357") E C ∈ {MATRICULADO, UPGRADE}.
      const dados = chunk
        .filter(function (l) {
          return _contemToken(l[cfg.FILTRO_COL_IDX], cfg.FILTRO_TXT) &&
                 cfg.FILTRO_STATUS_OK.includes(String(l[cfg.FILTRO_STATUS_IDX]).trim().toUpperCase());
        })
        .map(function (l) {
          return cfg.INDICES_COLUNAS.map(function (i) { return i < l.length ? l[i] : ''; });
        });

      if (dados.length > 0) {
        const proxLinha = abaDestino.getLastRow() + 1;
        abaDestino.getRange(proxLinha, 1, dados.length, dados[0].length).setValues(dados);
        totalImportado += dados.length;
      }

      linhaAtual += CHUNK_SIZE;
    }

    SpreadsheetApp.flush();
    _limparEstado(CHAVE_EST);
    _limparTriggersDe(NOME);

    const formulasDepois = _detectarFormulas(abaDestino);
    let msg = '✅ Subscritos 2020 importados!\n' +
              'Cabeçalho (linha ' + cfg.LINHA_CABECALHO + ') + ' + totalImportado +
              ' linha(s) com B = "' + cfg.FILTRO_TXT + '" e status ∈ {MATRICULADO, UPGRADE} em ' + cfg.ABA_DESTINO + '.';
    if (formulasAntes.length) {
      msg += '\n\n⚠️ CAUSA DO “SOME E VOLTA”: a aba continha FÓRMULA(S):\n' + formulasAntes.slice(0, 5).join('\n') +
             '\nUma fórmula que gera os dados (IMPORTRANGE/QUERY/ARRAYFORMULA/FILTER) sobrepõe o script. Foi removida.';
    }
    if (formulasDepois.length) {
      msg += '\n\n⛔ Fórmulas REAPARECERAM após a escrita: ' + formulasDepois.slice(0, 5).join(' | ') +
             '\nIndica GATILHO (onEdit/onChange) ou conector a reinseri-las.';
    }
    _avisar(msg);

  } catch (e) {
    _logErro(NOME, e);
    _avisar('❌ Erro: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
//  IMPORTAÇÃO 2 — DIRETORES (INCREMENTAL + AUTO-HIGIENIZAÇÃO)
//  R3: lock de execução.
//  Passo 0: remove duplicados existentes (D2: mantém a linha CONDECORADA).
//  D1 (decisão): mantém-se APPEND-ONLY — registros já existentes NÃO são
//  atualizados a partir da origem, para não sobrepor edições manuais no destino.
//  A antiga coluna REDE SOCIAL (col G) é apagada ao final (feature removida).
// =============================================================

function importarDiretores() {
  const NOME = 'importarDiretores';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    Logger.log('[' + NOME + '] Lock ocupado. Abortado.');
    return;
  }

  const cfg = CONFIG.DIRETORES;
  try {
    const abaOrigem  = _abrirAba(CONFIG.ID_ORIGEM, cfg.ABA_ORIGEM);
    const abaDestino = _abrirAbaLocal(cfg.ABA_DESTINO);

    // Passo 0 — deduplica o que já existe (idempotente).
    const removidosAntes = _removerDuplicadosPorId(abaDestino, cfg);

    const ultimaOrigem = abaOrigem.getLastRow();
    if (ultimaOrigem < cfg.SRC_LINHA_DADOS) {
      _avisar('ℹ️ Não há dados na aba DIRECTORS.' +
              (removidosAntes ? '\n(Duplicados removidos: ' + removidosAntes + ')' : ''));
      return;
    }

    const nLinhas = ultimaOrigem - cfg.SRC_LINHA_DADOS + 1;
    const matriz  = abaOrigem.getRange(cfg.SRC_LINHA_DADOS, 1, nLinhas, cfg.SRC_NUM_COLS).getValues();

    Logger.log('=== DIRETORES ===');
    Logger.log('getLastRow() origem = ' + ultimaOrigem + ' | linhas lidas = ' + matriz.length +
               ' | duplicados removidos = ' + removidosAntes);

    const largura = cfg.DST_NUM_COLS;

    const existentes = new Set();
    const ultimaDestino = abaDestino.getLastRow();
    if (ultimaDestino >= cfg.DST_LINHA_DADOS) {
      const chavesDst = abaDestino.getRange(cfg.DST_LINHA_DADOS, 1, ultimaDestino - cfg.DST_LINHA_DADOS + 1, 3).getValues();
      chavesDst.forEach(function (r) { existentes.add(_chaveDiretor(r[0], r[1], r[2])); });
    }

    const novos  = [];
    const vistos = new Set();
    let casaram = 0, pulados = 0;

    for (let i = 0; i < matriz.length; i++) {
      const orig = matriz[i];
      if (String(orig[cfg.FILTRO_COL_IDX]).indexOf(cfg.FILTRO_TXT) === -1) continue;
      casaram++;

      const chave = _chaveDiretor(orig[cfg.CHAVE_ID_SRC_IDX], orig[cfg.CHAVE_NOME_SRC_IDX], orig[cfg.CHAVE_NOMEN_SRC_IDX]);
      if (existentes.has(chave) || vistos.has(chave)) { pulados++; continue; }

      const linha = new Array(largura).fill('');
      cfg.MAP.forEach(function (m) { linha[m.dst - 1] = orig[m.src]; });
      novos.push(linha);
      vistos.add(chave);
    }

    Logger.log('Casaram: ' + casaram + ' | já existiam: ' + pulados + ' | novos: ' + novos.length);

    if (novos.length > 0) {
      const proxLinha = Math.max(abaDestino.getLastRow() + 1, cfg.DST_LINHA_DADOS);
      abaDestino.getRange(proxLinha, 1, novos.length, largura).setValues(novos);
      SpreadsheetApp.flush();
    }

    const redeSocialApagada = _limparRedeSocial(abaDestino);

    _avisar('✅ Diretores: ' + novos.length + ' novo(s) adicionado(s).\n' +
            'Duplicados removidos: ' + removidosAntes + '.\n' +
            casaram + ' casaram com o filtro, ' + pulados + ' já existiam.\n' +
            'Rede social (col G) apagada em ' + redeSocialApagada + ' linha(s).');

  } catch (e) {
    _logErro(NOME, e);
    _avisar('❌ Erro: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
//  LIMPEZA AVULSA — REMOVER DIRETORES DUPLICADOS
// =============================================================

function limparDuplicadosDiretores() {
  const NOME = 'limparDuplicadosDiretores';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) {
    Logger.log('[' + NOME + '] Lock ocupado. Abortado.');
    return;
  }

  const cfg = CONFIG.DIRETORES;
  try {
    const abaDestino        = _abrirAbaLocal(cfg.ABA_DESTINO);
    const removidos         = _removerDuplicadosPorId(abaDestino, cfg);
    const redeSocialApagada = _limparRedeSocial(abaDestino);
    SpreadsheetApp.flush();
    _avisar('🧹 Limpeza concluída.\n' +
            'Diretores duplicados removidos: ' + removidos + '.\n' +
            'Rede social (col G) apagada em ' + redeSocialApagada + ' linha(s).');
  } catch (e) {
    _logErro(NOME, e);
    _avisar('❌ Erro: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
//  AUXILIARES PRIVADAS
// =============================================================

/**
 * D3: verdadeiro se `valor`, separado por não-dígitos, contém `token` como
 * número inteiro isolado. Aceita "7388" e "7388 / 8357"; rejeita "73880".
 */
function _contemToken(valor, token) {
  const partes = String(valor == null ? '' : valor).split(/[^0-9]+/);
  return partes.indexOf(String(token)) !== -1;
}

/**
 * D2: pontua uma linha do DESTINO para escolher, entre duplicados do mesmo ID,
 * a mais completa: condecorada (col E = TRUE) vale mais; ter DATA (col F) desempata.
 */
function _scoreDiretor(r, cfg) {
  const cond = r[cfg.DST_COND_IDX];
  const condFlag = (cond === true) || (String(cond).trim().toUpperCase() === 'TRUE');
  const temData  = String(r[cfg.DST_DATA_IDX] == null ? '' : r[cfg.DST_DATA_IDX]).trim() !== '';
  let s = 0;
  if (condFlag) s += 2;
  if (temData)  s += 1;
  return s;
}

/** Lista fórmulas presentes numa aba (até 10), em A1 → fórmula. Diagnóstico do "revert". */
function _detectarFormulas(aba) {
  const rng = aba.getDataRange();
  if (rng.getNumRows() === 0 || rng.getNumColumns() === 0) return [];
  const f = rng.getFormulas();
  const achados = [];
  for (let r = 0; r < f.length; r++) {
    for (let c = 0; c < f[r].length; c++) {
      if (f[r][c] !== '') {
        achados.push(rng.getCell(r + 1, c + 1).getA1Notation() + ' -> ' + f[r][c]);
        if (achados.length >= 10) return achados;
      }
    }
  }
  return achados;
}

/** Remove gatilhos de RELÓGIO (continuação) de uma função. Retorna quantos. */
function _limparTriggersDe(nomeFuncao) {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === nomeFuncao && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

/**
 * D2: remove duplicados em DIRETORES por chave (_chaveDiretor), mantendo a linha
 * de MAIOR score (condecorada). Preserva a ordem da 1ª aparição de cada chave.
 * Retorna o número de linhas removidas.
 */
function _removerDuplicadosPorId(abaDestino, cfg) {
  const primeira = cfg.DST_LINHA_DADOS;
  const ultima   = abaDestino.getLastRow();
  if (ultima < primeira) return 0;

  const n       = ultima - primeira + 1;
  const largura = cfg.DST_NUM_COLS;
  const dados   = abaDestino.getRange(primeira, 1, n, largura).getValues();

  const ordem  = [];   // chaves na ordem de 1ª aparição
  const melhor = {};   // chave -> { row, score }

  dados.forEach(function (r) {
    const chave = _chaveDiretor(r[0], r[1], r[2]);
    const sc = _scoreDiretor(r, cfg);
    if (!(chave in melhor)) {
      melhor[chave] = { row: r, score: sc };
      ordem.push(chave);
    } else if (sc > melhor[chave].score) {
      melhor[chave].row = r;       // troca pela mais completa; ordem mantém-se
      melhor[chave].score = sc;
    }
  });

  const mantidos  = ordem.map(function (k) { return melhor[k].row; });
  const removidos = n - mantidos.length;
  if (removidos <= 0) return 0;

  abaDestino.getRange(primeira, 1, mantidos.length, largura).setValues(mantidos);
  abaDestino.deleteRows(primeira + mantidos.length, removidos);

  Logger.log('[DEDUP] Removidos: ' + removidos + ' | Mantidos: ' + mantidos.length);
  return removidos;
}

/**
 * Apaga por completo a antiga coluna REDE SOCIAL (col G) da aba DIRETORES,
 * incluindo o cabeçalho. Idempotente. Retorna quantas linhas foram limpas.
 */
function _limparRedeSocial(abaDestino) {
  const COL_G  = 7;   // col G (antiga REDE SOCIAL), 1-based
  const ultima = abaDestino.getLastRow();
  if (ultima < 1) return 0;
  abaDestino.getRange(1, COL_G, ultima, 1).clearContent();
  SpreadsheetApp.flush();
  Logger.log('[REDE SOCIAL] Coluna G apagada em ' + ultima + ' linha(s).');
  return ultima;
}

/** Chave de identidade: usa o ID; se vazio, cai para NOME|NOMENCLATURA. */
function _chaveDiretor(id, nome, nomenclatura) {
  const i = String(id == null ? '' : id).trim();
  if (i !== '') return 'ID:' + i;
  const n = String(nome         == null ? '' : nome).trim().toUpperCase();
  const c = String(nomenclatura == null ? '' : nomenclatura).trim().toUpperCase();
  return 'NM:' + n + '|' + c;
}

/** Mostra alerta só quando há contexto de UI (menu); silencioso em gatilhos. */
function _avisar(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); }
  catch (e) { Logger.log(mensagem); }
}

function _abrirAba(idPlanilha, nomeAba) {
  let planilha;
  try { planilha = SpreadsheetApp.openById(idPlanilha); }
  catch (e) {
    throw new Error('Não foi possível abrir a planilha de origem.\n' +
      'Verifique: 1) ID correto: ' + idPlanilha + ' 2) permissão de leitura 3) planilha não apagada.');
  }
  const aba = planilha.getSheetByName(nomeAba);
  if (!aba) throw new Error("Aba '" + nomeAba + "' não encontrada na origem (" + idPlanilha + ").");
  return aba;
}

function _abrirAbaLocal(nomeAba) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  if (!aba) throw new Error("Aba de destino '" + nomeAba + "' não encontrada nesta planilha.");
  return aba;
}

function _abrirOuCriarAbaLocal(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(nomeAba);
  if (!aba) aba = ss.insertSheet(nomeAba);
  return aba;
}

function _agendarContinuacao(nomeFuncao) {
  ScriptApp.getProjectTriggers()
    .filter(function (t) {
      return t.getHandlerFunction() === nomeFuncao && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK;
    })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(nomeFuncao).timeBased().after(60 * 1000).create();
}

function _salvarEstado(chave, estado) {
  PropertiesService.getScriptProperties().setProperty('ESTADO_' + chave, JSON.stringify(estado));
}

function _carregarEstado(chave) {
  const val = PropertiesService.getScriptProperties().getProperty('ESTADO_' + chave);
  return val ? JSON.parse(val) : {};
}

function _limparEstado(chave) {
  PropertiesService.getScriptProperties().deleteProperty('ESTADO_' + chave);
}

function _logErro(nomeFuncao, erro) {
  Logger.log('[ERRO] ' + nomeFuncao + ': ' + erro.message + '\n' + erro.stack);
}
