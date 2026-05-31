// Emissão de petição inicial em DOCX e PDF a partir dos templates Word oficiais
// do escritório Andrade & Fukushima.
//
// REGRA ABSOLUTA: este módulo NÃO recria, NÃO reformata e NÃO reconstrói nada.
// Apenas abre o .docx original, substitui placeholders {KEY} pelos valores
// preenchidos e devolve o mesmo arquivo. Toda a diagramação (negritos,
// itálicos, tabelas, margens, fontes, cabeçalhos e rodapés) é herdada
// integralmente do template Word.
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
// file-saver é CommonJS — usa default import para evitar erro em SSR
import FileSaver from "file-saver";
const { saveAs } = FileSaver;

const TEMPLATES = {
  "acima-bacen": "/PETICAO_INICIAL_-_ACIMA_DO_BACEN.docx",
  "abaixo-bacen": "/PETICAO_INICIAL_-_ABAIXO_DO_BACEN.docx",
};

// Placeholders presentes no template Word do escritório. Usados para validar
// que nenhum campo ficou sem preenchimento antes da emissão.
const TEMPLATE_KEYS_COMUNS = [
  "NOME", "NACIONALIDADE", "ESTADO_CIVIL", "PROFISSAO", "RG", "CPF", "ENDERECO",
  "BANCO", "CNPJ_BANCO", "ENDERECO_BANCO",
  "CIDADE", "UF", "DATA", "DATA_CONTRATO",
  "VEICULO", "ANO_VEICULO", "NUMERO_CONTRATO",
  "VALOR_PRINCIPAL", "VALOR_LIQUIDO", "VALOR_TOTAL", "VALOR_PARCELA",
  "QTD_PARCELAS", "TAXA_JUROS_MES", "TAXA_JUROS_ANO", "CET",
  "TARIFA_CADASTRO", "TARIFA_AVALIACAO", "DESPESAS_REGISTRO",
  "SEGURO", "SEGURADORA", "CNPJ_SEGURADORA",
  "VALOR_CAUSA",
  "JURIS_TARIFA_AVALIACAO", "JURIS_TARIFA_REGISTRO", "JURIS_SEGURO",
];
const TEMPLATE_KEYS_POR_TESE = {
  "acima-bacen":  [...TEMPLATE_KEYS_COMUNS, "JURIS_ACIMA_BACEN", "TAXA_BACEN_MES", "TAXA_BACEN_ANO"],
  "abaixo-bacen": [...TEMPLATE_KEYS_COMUNS],
};

const cache = {};
async function loadTemplate(tese) {
  const url = TEMPLATES[tese];
  if (!url) throw new Error(`Tese inválida: ${tese}`);
  if (cache[tese]) return cache[tese];
  const r = await fetch(url, { cache: "no-store" });
  if (!r.ok) throw new Error(`Falha ao carregar template (${tese}): HTTP ${r.status}`);
  cache[tese] = await r.arrayBuffer();
  return cache[tese];
}

// Mapeia os campos do formulário e as jurisprudências do estado selecionado
// para as chaves exatas do template Word.
export function buildData(form, juris) {
  const j = (juris && juris[form.uf]) || {};
  const v = (x) => (x === null || x === undefined ? "" : String(x).trim());
  return {
    // Qualificação
    NOME: v(form.nome),
    NACIONALIDADE: v(form.nacionalidade),
    ESTADO_CIVIL: v(form.estadoCivil),
    PROFISSAO: v(form.profissao),
    RG: v(form.rg),
    CPF: v(form.cpf),
    ENDERECO: v(form.endereco),
    // Banco
    BANCO: v(form.banco),
    CNPJ_BANCO: v(form.cnpjBanco),
    ENDERECO_BANCO: v(form.enderecoBanco),
    // Local / data
    CIDADE: v(form.cidade),
    UF: v(form.uf),
    DATA: v(form.data),
    DATA_CONTRATO: v(form.dataContrato),
    // Veículo / contrato
    VEICULO: v(form.veiculo),
    ANO_VEICULO: v(form.anoVeiculo),
    NUMERO_CONTRATO: v(form.numeroContrato),
    // Valores
    VALOR_PRINCIPAL: v(form.valorPrincipal),
    VALOR_LIQUIDO: v(form.valorLiquido),
    VALOR_TOTAL: v(form.valorTotal),
    VALOR_PARCELA: v(form.valorParcela),
    QTD_PARCELAS: v(form.qtdParcelas),
    // Taxas do contrato
    TAXA_JUROS_MES: v(form.taxaJurosContrMes || form.taxaJurosMes),
    TAXA_JUROS_ANO: v(form.taxaJurosContrAno || form.taxaJurosAno),
    CET: v(form.cet),
    // Encargos
    TARIFA_CADASTRO: v(form.tarifaCadastro),
    TARIFA_AVALIACAO: v(form.tarifaAvaliacao),
    DESPESAS_REGISTRO: v(form.despesasRegistro),
    SEGURO: v(form.seguro),
    SEGURADORA: v(form.seguradora),
    CNPJ_SEGURADORA: v(form.cnpjSeguradora),
    // Tese acima do BACEN
    TAXA_BACEN_MES: v(form.taxaBacenMes),
    TAXA_BACEN_ANO: v(form.taxaBacenAno),
    // Causa
    VALOR_CAUSA: v(form.valorCausa),
    // Jurisprudências do estado selecionado
    JURIS_ACIMA_BACEN: v(j.acimaBacen),
    JURIS_TARIFA_AVALIACAO: v(j.tarifaAvaliacao),
    JURIS_TARIFA_REGISTRO: v(j.tarifaRegistro),
    JURIS_SEGURO: v(j.seguro),
    // Flags condicionais (checkboxes). Default = true.
    incluir_seguro:           form.incluir_seguro           !== false,
    incluir_tarifa_cadastro:  form.incluir_tarifa_cadastro  !== false,
    incluir_tarifa_avaliacao: form.incluir_tarifa_avaliacao !== false,
    incluir_tarifa_registro:  form.incluir_tarifa_registro  !== false,
    incluir_cet:              form.incluir_cet              !== false,
  };
}

// Mapa amigável de chave do template -> rótulo legível.
const LABEL = {
  NOME: "Nome do cliente", NACIONALIDADE: "Nacionalidade", ESTADO_CIVIL: "Estado civil",
  PROFISSAO: "Profissão", RG: "RG", CPF: "CPF", ENDERECO: "Endereço",
  BANCO: "Banco réu", CNPJ_BANCO: "CNPJ do banco", ENDERECO_BANCO: "Endereço do banco",
  CIDADE: "Cidade", UF: "UF", DATA: "Data", DATA_CONTRATO: "Data do contrato",
  VEICULO: "Veículo financiado", ANO_VEICULO: "Ano do veículo",
  NUMERO_CONTRATO: "Nº do contrato",
  VALOR_PRINCIPAL: "Valor principal", VALOR_LIQUIDO: "Valor líquido",
  VALOR_TOTAL: "Valor total", VALOR_PARCELA: "Valor da parcela",
  QTD_PARCELAS: "Quantidade de parcelas",
  TAXA_JUROS_MES: "Taxa de juros (mês)", TAXA_JUROS_ANO: "Taxa de juros (ano)",
  CET: "CET", TARIFA_CADASTRO: "Tarifa de cadastro",
  TARIFA_AVALIACAO: "Tarifa de avaliação", DESPESAS_REGISTRO: "Despesas de registro",
  SEGURO: "Seguro", SEGURADORA: "Seguradora", CNPJ_SEGURADORA: "CNPJ da seguradora",
  VALOR_CAUSA: "Valor da causa",
  TAXA_BACEN_MES: "Taxa BACEN (mês)", TAXA_BACEN_ANO: "Taxa BACEN (ano)",
  JURIS_ACIMA_BACEN: "Jurisprudência – taxa acima do BACEN",
  JURIS_TARIFA_AVALIACAO: "Jurisprudência – tarifa de avaliação",
  JURIS_TARIFA_REGISTRO: "Jurisprudência – despesas de registro",
  JURIS_SEGURO: "Jurisprudência – seguro prestamista",
};

// Mapeia chaves de campo para o checkbox que as torna opcionais.
// Se o checkbox estiver desmarcado, o bloco é removido do DOCX e o campo
// correspondente deixa de ser obrigatório.
const KEY_REQUIRES_FLAG = {
  TARIFA_CADASTRO:        "incluir_tarifa_cadastro",
  TARIFA_AVALIACAO:       "incluir_tarifa_avaliacao",
  DESPESAS_REGISTRO:      "incluir_tarifa_registro",
  SEGURO:                 "incluir_seguro",
  SEGURADORA:             "incluir_seguro",
  CNPJ_SEGURADORA:        "incluir_seguro",
  CET:                    "incluir_cet",
  JURIS_TARIFA_AVALIACAO: "incluir_tarifa_avaliacao",
  JURIS_TARIFA_REGISTRO:  "incluir_tarifa_registro",
  JURIS_SEGURO:           "incluir_seguro",
};

export function validateData(form, juris) {
  const expected = TEMPLATE_KEYS_POR_TESE[form.tese] || TEMPLATE_KEYS_COMUNS;
  const data = buildData(form, juris);
  const missing = expected.filter((k) => {
    if (data[k]) return false;
    const flag = KEY_REQUIRES_FLAG[k];
    if (flag && data[flag] === false) return false; // bloco desligado → opcional
    return true;
  });
  return missing.map((k) => LABEL[k] || k);
}

async function renderDocx(form, juris) {
  const missing = validateData(form, juris);
  if (missing.length) {
    throw new Error(
      "Campos faltantes para emitir a petição:\n• " + missing.join("\n• ")
    );
  }
  const buf = await loadTemplate(form.tese);
  const zip = new PizZip(buf);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{", end: "}" },
    nullGetter: () => "",
  });
  doc.render(buildData(form, juris));
  return doc.getZip().generate({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}

function fileBase(form) {
  const nome = (form.nome || "Peticao_Inicial").trim().replace(/\s+/g, "_");
  return `Peticao_Inicial_${nome}`;
}

export async function emitirDOCX(form, juris) {
  const blob = await renderDocx(form, juris);
  saveAs(blob, `${fileBase(form)}.docx`);
}

// PDF: renderiza o .docx gerado com docx-preview em uma nova janela e dispara
// o diálogo de impressão (Salvar como PDF). Preserva fonte, tabelas,
// cabeçalhos, rodapés e paginação exatamente como definidos no Word.
export async function emitirPDF(form, juris) {
  const blob = await renderDocx(form, juris);
  const { renderAsync } = await import("docx-preview");

  const win = window.open("", "_blank");
  if (!win) throw new Error("Pop-ups bloqueados. Permita pop-ups para gerar o PDF.");
  win.document.write(`<!doctype html><html><head>
    <meta charset="utf-8"/>
    <title>${fileBase(form)}</title>
    <style>
      body { margin: 0; background: #fff; }
      .docx-wrapper { background: #fff !important; padding: 0 !important; }
      .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 auto !important; }
      @media print {
        .docx-wrapper > section.docx { margin: 0 !important; box-shadow: none !important; }
        @page { margin: 0; }
      }
    </style>
  </head><body><div id="container"></div></body></html>`);
  win.document.close();

  const container = win.document.getElementById("container");
  await renderAsync(blob, container, null, {
    className: "docx",
    inWrapper: true,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    experimental: true,
    useBase64URL: true,
  });

  if (win.document.fonts && win.document.fonts.ready) {
    try { await win.document.fonts.ready; } catch (e) { /* noop */ }
  }
  await new Promise((r) => setTimeout(r, 400));
  win.focus();
  win.print();
}
