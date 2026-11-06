// Carrega dados do dashboard a partir de um arquivo JavaScript externo.
// Defina window.LABOR_DASH_SCRIPT_URL antes deste arquivo para usar SharePoint/URL absoluta.
//
// Aceita dois formatos de arquivo:
// - Formato pronto: o arquivo ja define window.LABOR_DASH diretamente.
// - Formato bruto (Power Automate): o arquivo define window.LABOR_DASH_RAW_ROWS,
//   com uma linha por cliente e os nomes de coluna originais da consulta DAX
//   (ex: "d_Cliente[ID Cliente]", "[Fat. Media Janela 1 Coligada]"). Esse loader
//   transforma esse formato bruto no window.LABOR_DASH que o dashboard espera.
(function(){
  // O support.js reinsere os scripts do <head>; sem esta trava o loader rodaria duas
  // vezes e a segunda rodada apagaria os dados ja carregados pela primeira.
  if(window.LABOR_DASH_LOADER_STARTED) return;
  window.LABOR_DASH_LOADER_STARTED = true;
  const DEFAULT_URL = './labor-data.js';
  const url = window.LABOR_DASH_SCRIPT_URL || DEFAULT_URL;

  function num(v){
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  function janelas(row, prefix){
    return Array.from({length:7}, (_, i) => num(row[prefix.replace('{n}', i + 1)]));
  }

  function transformRawRows(rawRows){
    const RAW = rawRows.map(row => {
      const fms = janelas(row, '[Fat. Media Janela {n} Coligada]');
      const mgs = janelas(row, '[Margem Media Janela {n} Coligada]');
      const mcs = janelas(row, '[Meses Compra Janela {n} Coligada]');
      const sks = janelas(row, '[Media SKUs Janela {n}]');
      const mbs = Array.from({length:7}, (_, i) => num(row['[Pct_Margem_Janela_' + (i + 1) + '_Coligada]']) * 100);
      return {
        id: String(row['d_Cliente[ID Cliente]'] || '').trim(),
        col: String(row['d_Cliente[Coligada]'] || '').trim(),
        cnpj: String(row['d_Cliente[CNPJ Cliente]'] || '').trim(),
        no: String(row['d_Cliente[Nome Fantasia]'] || '').trim(),
        vd: String(row['d_Cliente[Nome Vendedor Proper]'] || '').trim(),
        fm: fms[0], fm7: fms[6], fms,
        mg: mgs[0], mg7: mgs[6], mgs,
        mc: mcs[0], m7: mcs[6], mcs,
        sk: sks[0], sk7: sks[6], sks,
        mb: mbs[0], mb7: mbs[6], mbs
      };
    });
    return { COUNTS: { auto: {}, total: RAW.length }, RAW };
  }

  // Sem login na Microsoft ou sem permissao no SharePoint, o navegador pode ficar
  // preso no redirecionamento de login sem disparar onerror. O timeout garante que
  // o aviso de acesso apareca mesmo assim.
  const TIMEOUT_MS = 20000;
  const PROBE_TIMEOUT_MS = 10000;
  const STORAGE_KEY = 'labor_dash_carteira';
  const DATA_GLOBALS = ['LABOR_DASH', 'LABOR_DASH_RAW_ROWS', 'LABOR_DASH_META', 'LABOR_DASH_CARTEIRAS'];

  function comVersao(u){ return u + (u.includes('?') ? '&' : '?') + 'v=' + Date.now(); }

  // Carrega um .js do SharePoint e devolve as variaveis que ele criou, limpando-as
  // do window logo em seguida. Assim varios arquivos podem ser testados ao mesmo
  // tempo sem um sobrescrever o outro. Arquivo sem permissao: o SharePoint devolve
  // erro ou uma pagina HTML (que nao cria nenhuma variavel) e a promessa falha.
  function capturarScript(u, timeoutMs){
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      const timer = setTimeout(() => { script.remove(); reject(new Error('timeout')); }, timeoutMs);
      script.src = comVersao(u);
      script.async = true;
      script.onload = () => {
        clearTimeout(timer);
        const got = {};
        DATA_GLOBALS.forEach(k => {
          got[k] = window[k];
          try { delete window[k]; } catch(e){}
          if(k in window) window[k] = undefined;
        });
        script.remove();
        resolve(got);
      };
      script.onerror = () => { clearTimeout(timer); script.remove(); reject(new Error('sem acesso')); };
      document.head.appendChild(script);
    });
  }

  function temDados(got){
    return !!(got && ((got.LABOR_DASH && got.LABOR_DASH.RAW) || Array.isArray(got.LABOR_DASH_RAW_ROWS)));
  }

  function aplicarDados(got, origem){
    if(got.LABOR_DASH && got.LABOR_DASH.RAW){
      window.LABOR_DASH = got.LABOR_DASH;
    } else {
      window.LABOR_DASH_RAW_ROWS = got.LABOR_DASH_RAW_ROWS;
      window.LABOR_DASH = transformRawRows(got.LABOR_DASH_RAW_ROWS);
    }
    window.LABOR_DASH_META = got.LABOR_DASH_META || { source: 'Dados: ' + origem };
    return window.LABOR_DASH;
  }

  function lerSalvo(){ try { return localStorage.getItem(STORAGE_KEY) || ''; } catch(e){ return ''; } }
  function salvar(v){ try { localStorage.setItem(STORAGE_KEY, v); } catch(e){} }
  function esquecer(){ try { localStorage.removeItem(STORAGE_KEY); } catch(e){} }
  const CARTEIRA_OK = /^(vendedores|grupos)\/[a-z0-9_-]+$/i;

  const auto = window.LABOR_DASH_AUTO || null;
  const msgSemAcesso = 'Nao foi possivel carregar os dados do SharePoint. Entre com sua conta Microsoft/SharePoint em outra aba, confirme que voce tem permissao e recarregue este dashboard.';

  // Link unico (sem ?carteira=): descobre sozinho qual arquivo a pessoa pode abrir.
  // 1) Gestao (labor-data.js) e a carteira lembrada neste navegador, em paralelo;
  // 2) se nenhum abrir, le carteiras/indice.js e testa todas as carteiras ao mesmo
  //    tempo, ficando com a primeira que o SharePoint liberar (grupos antes de vendedores).
  // O sigilo continua nas permissoes do SharePoint: so abre o que a pessoa ja podia abrir.
  async function descobrir(){
    const urlCarteira = c => auto.base + c + '.js';
    const escolher = (got, carteira) => {
      window.LABOR_DASH_CARTEIRA = carteira || undefined;
      window.LABOR_DASH_OPERACIONAL_SCRIPT_URL = carteira ? auto.base + carteira + '-operacional.js' : auto.gestaoOperacionalUrl;
      salvar(carteira || 'gestao');
      return aplicarDados(got, carteira ? urlCarteira(carteira) : auto.gestaoUrl);
    };
    const salvo = lerSalvo();
    const tentativas = [capturarScript(auto.gestaoUrl, TIMEOUT_MS).then(g => temDados(g) ? g : null, () => null)];
    if(CARTEIRA_OK.test(salvo)) tentativas.push(capturarScript(urlCarteira(salvo), PROBE_TIMEOUT_MS).then(g => temDados(g) ? g : null, () => null));
    const [gestao, lembrada] = await Promise.all(tentativas);
    if(gestao) return escolher(gestao, '');
    if(lembrada) return escolher(lembrada, salvo);

    const idx = await capturarScript(auto.indiceUrl, PROBE_TIMEOUT_MS).catch(() => null);
    const lista = (idx && Array.isArray(idx.LABOR_DASH_CARTEIRAS) ? idx.LABOR_DASH_CARTEIRAS : [])
      .map(c => String(c || '').trim()).filter(c => CARTEIRA_OK.test(c) && c !== salvo);
    const ordem = [...new Set([...lista.filter(c => c.startsWith('grupos/')), ...lista.filter(c => c.startsWith('vendedores/'))])];
    const resultados = await Promise.all(ordem.map(c =>
      capturarScript(urlCarteira(c), PROBE_TIMEOUT_MS).then(g => temDados(g) ? g : null, () => null)));
    const i = resultados.findIndex(Boolean);
    if(i >= 0) return escolher(resultados[i], ordem[i]);
    esquecer();
    throw new Error(msgSemAcesso);
  }

  window.LABOR_DASH_READY = auto ? descobrir() : capturarScript(url, TIMEOUT_MS).then(got => {
    if(temDados(got)) return aplicarDados(got, url);
    throw new Error('Arquivo de dados carregou, mas nao criou window.LABOR_DASH nem window.LABOR_DASH_RAW_ROWS. Link dos dados: ' + url);
  }, err => {
    throw new Error(err.message === 'timeout'
      ? 'Tempo esgotado ao carregar os dados do SharePoint. Link dos dados: ' + url
      : msgSemAcesso + ' Link dos dados: ' + url);
  });

  // Dados operacionais (prev` opcional: previa do mes em andamento, calculada com
  // janela dinamica ate hoje). So carrega se a pagina definir a URL; se nao definir,
  // resolve como null sem gerar erro (recurso opcional, nao trava o dashboard oficial).

  function transformOperacionalRows(rawRows){
    const byId = {};
    rawRows.forEach(row => {
      const id = String(row['d_Cliente[ID Cliente]'] || '').trim();
      if(!id) return;
      const fm = num(row['[Faturamento_Operacional]']);
      const mg = num(row['[Margem_Operacional]']);
      byId[id] = {
        id,
        col: String(row['d_Cliente[Coligada]'] || '').trim(),
        cnpj: String(row['d_Cliente[CNPJ Cliente]'] || '').trim(),
        no: String(row['d_Cliente[Nome Fantasia]'] || '').trim(),
        vd: String(row['d_Cliente[Nome Vendedor Proper]'] || '').trim(),
        fm, mg,
        mb: fm > 0 ? (mg / fm) * 100 : 0,
        sk: num(row['[SKU_Operacional]']),
        mc: num(row['[Meses_Operacional]'])
      };
    });
    return { byId };
  }

  // A URL operacional so e conhecida depois que a carteira e descoberta (link unico),
  // por isso espera o carregamento principal.
  window.LABOR_DASH_OPERACIONAL_READY = window.LABOR_DASH_READY.then(() => {
    const operacionalUrl = window.LABOR_DASH_OPERACIONAL_SCRIPT_URL || null;
    if(!operacionalUrl) return null;
    return new Promise((resolve) => {
      const script = document.createElement('script');
      script.src = comVersao(operacionalUrl);
      script.async = true;
      script.onload = () => {
        if(Array.isArray(window.LABOR_DASH_OPERACIONAL_ROWS)){
          window.LABOR_DASH_OPERACIONAL = transformOperacionalRows(window.LABOR_DASH_OPERACIONAL_ROWS);
        }
        resolve(window.LABOR_DASH_OPERACIONAL || null);
      };
      script.onerror = () => resolve(null);
      document.head.appendChild(script);
    });
  }, () => null);
})();
