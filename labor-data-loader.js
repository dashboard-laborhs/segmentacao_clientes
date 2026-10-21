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
  const DEFAULT_URL = './labor-data.js';
  const url = window.LABOR_DASH_SCRIPT_URL || DEFAULT_URL;
  const scriptUrl = url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now();

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

  window.LABOR_DASH_READY = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('Tempo esgotado ao carregar os dados do SharePoint. Link dos dados: ' + url));
    }, TIMEOUT_MS);
    const script = document.createElement('script');
    script.src = scriptUrl;
    script.async = true;
    script.onload = () => {
      clearTimeout(timer);
      if(window.LABOR_DASH && window.LABOR_DASH.RAW){
        if(!window.LABOR_DASH_META) window.LABOR_DASH_META = { source: 'Dados: ' + url };
        resolve(window.LABOR_DASH);
        return;
      }
      if(Array.isArray(window.LABOR_DASH_RAW_ROWS)){
        window.LABOR_DASH = transformRawRows(window.LABOR_DASH_RAW_ROWS);
        if(!window.LABOR_DASH_META) window.LABOR_DASH_META = { source: 'Dados: ' + url };
        resolve(window.LABOR_DASH);
        return;
      }
      reject(new Error('Arquivo de dados carregou, mas nao criou window.LABOR_DASH nem window.LABOR_DASH_RAW_ROWS.'));
    };
    script.onerror = () => {
      clearTimeout(timer);
      reject(new Error('Nao foi possivel carregar os dados do SharePoint. Entre com sua conta Microsoft/SharePoint em outra aba, confirme que voce tem permissao para o arquivo labor-data.js e recarregue este dashboard. Link dos dados: ' + url));
    };
    document.head.appendChild(script);
  });

  // Dados operacionais (prev` opcional: previa do mes em andamento, calculada com
  // janela dinamica ate hoje). So carrega se a pagina definir a URL; se nao definir,
  // resolve como null sem gerar erro (recurso opcional, nao trava o dashboard oficial).
  const operacionalUrl = window.LABOR_DASH_OPERACIONAL_SCRIPT_URL || null;

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

  window.LABOR_DASH_OPERACIONAL_READY = operacionalUrl ? new Promise((resolve) => {
    const opScriptUrl = operacionalUrl + (operacionalUrl.includes('?') ? '&' : '?') + 'v=' + Date.now();
    const script = document.createElement('script');
    script.src = opScriptUrl;
    script.async = true;
    script.onload = () => {
      if(Array.isArray(window.LABOR_DASH_OPERACIONAL_ROWS)){
        window.LABOR_DASH_OPERACIONAL = transformOperacionalRows(window.LABOR_DASH_OPERACIONAL_ROWS);
      }
      resolve(window.LABOR_DASH_OPERACIONAL || null);
    };
    script.onerror = () => resolve(null);
    document.head.appendChild(script);
  }) : Promise.resolve(null);
})();
