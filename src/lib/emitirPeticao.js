// Emissão de petição inicial em DOCX e PDF a partir do template Word
// do escritório Andrade & Fukushima (preserva 100% da diagramação original).
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

const TEMPLATE_URL = "/template-peticao.docx";

let templateCache = null;
async function loadTemplate() {
  if (templateCache) return templateCache;
  const resp = await fetch(TEMPLATE_URL);
  if (!resp.ok) throw new Error("Não foi possível carregar o template Word.");
  templateCache = await resp.arrayBuffer();
  return templateCache;
}

// Mapeia o estado do formulário + jurisprudências do estado para os placeholders
// presentes no template (.docx pré-processado: tags no formato {CHAVE}).
export function buildData(form, juris) {
  const j = (juris && juris[form.uf]) || {};
  const fb = (v, alt) => (v && String(v).trim() ? String(v) : alt || "______");
  return {
    // Qualificação
    NOME: fb(form.nome, "[NOME]"),
    RG: fb(form.rg, "[RG]"),
    CPF: fb(form.cpf, "[CPF]"),
    ENDERECO: fb(form.endereco, "[ENDEREÇO]"),
    BANCO: fb(form.banco, "[BANCO]"),
    CNPJ_BANCO: fb(form.cnpjBanco, "[CNPJ]"),
    ENDERECO_BANCO: fb(form.enderecoBanco, "[ENDEREÇO DO BANCO]"),
    CIDADE: fb(form.cidade, "[CIDADE]"),
    UF: fb(form.uf, "[UF]"),
    DATA: fb(form.data, "[DATA]"),
    DATA_CONTRATO: fb(form.dataContrato, "[DATA DO CONTRATO]"),
    VEICULO: fb(form.veiculo, "[CARRO FINANCIADO]"),
    ANO_VEICULO: fb(form.anoVeiculo, "[ANO]"),
    VALOR_CAUSA: fb(form.valorCausa, "[VALOR DA CAUSA]"),
    // Tabela – Dados essenciais do contrato
    NUMERO_CONTRATO: fb(form.numeroContrato),
    VALOR_PRINCIPAL: fb(form.valorPrincipal),
    VALOR_LIQUIDO: fb(form.valorLiquido),
    VALOR_TOTAL: fb(form.valorTotal),
    TAXA_JUROS: fb(form.taxaJurosContrAno || form.taxaJurosContrMes),
    CET: fb(form.cet),
    QTD_PARCELAS: fb(form.qtdParcelas),
    VALOR_PARCELA: fb(form.valorParcela),
    TARIFA_AVALIACAO_TBL: fb(form.tarifaAvaliacao),
    DESPESAS_REGISTRO: fb(form.despesasRegistro),
    SEGURO: fb(form.seguro),
    TARIFA_CADASTRO_TBL: fb(form.tarifaCadastro),
    TAXA_JUROS_MES: fb(form.taxaJurosContrMes || form.taxaJurosMes),
    TAXA_JUROS_ANO: fb(form.taxaJurosContrAno || form.taxaJurosAno),
    CNPJ_SEGURADORA: fb(form.cnpjSeguradora),
    // Encargos isolados
    TARIFA_AVALIACAO: fb(form.tarifaAvaliacao),
    TARIFA_CADASTRO: fb(form.tarifaCadastro),
    TAXA_BACEN: fb(form.taxaBacenAno || form.taxaBacenMes),
    SEGURADORA: fb(form.seguradora, "[SEGURADORA]"),
    // Jurisprudências do estado selecionado
    JURIS_ACIMA_BACEN: fb(j.acimaBacen, "[Adicionar jurisprudência local]"),
    JURIS_TARIFA_AVALIACAO: fb(j.tarifaAvaliacao, "[Adicionar jurisprudência local]"),
    JURIS_TARIFA_REGISTRO: fb(j.tarifaRegistro, "[Adicionar jurisprudência local]"),
    JURIS_SEGURO: fb(j.seguro, "[Adicionar jurisprudência local]"),
  };
}

async function renderDocx(form, juris) {
  const buf = await loadTemplate();
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

// PDF: renderiza o .docx gerado com docx-preview em uma nova janela e
// dispara o diálogo de impressão (Salvar como PDF). Preserva fielmente
// fonte Arial, tabelas, cabeçalhos, rodapés e paginação do Word.
export async function emitirPDF(form, juris) {
  const blob = await renderDocx(form, juris);
  const { renderAsync } = await import("docx-preview");

  const win = window.open("", "_blank");
  if (!win) throw new Error("Pop-ups bloqueados. Permita pop-ups para gerar o PDF.");
  win.document.write(`<!doctype html><html><head>
    <meta charset="utf-8"/>
    <title>${fileBase(form)}</title>
    <style>
      body { margin: 0; background: #fff; font-family: Arial, sans-serif; }
      .docx-wrapper { background: #fff !important; padding: 0 !important; }
      .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 auto !important; }
      @media print {
        .docx-wrapper > section.docx { margin: 0 !important; box-shadow: none !important; }
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
    breakPages: true,
    experimental: true,
    useBase64URL: true,
  });

  // Aguarda fontes/imagens carregarem
  if (win.document.fonts && win.document.fonts.ready) {
    try { await win.document.fonts.ready; } catch (e) { /* noop */ }
  }
  await new Promise((r) => setTimeout(r, 300));
  win.focus();
  win.print();
}
