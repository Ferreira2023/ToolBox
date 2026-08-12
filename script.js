const app = document.getElementById("app");
const pageTitle = document.getElementById("pageTitle");
const menuButtons = document.querySelectorAll(".menu-btn");

async function carregarFerramenta(arquivo, titulo, botaoAtivo = null) {
  try {
    app.innerHTML = '<div class="loading-page">Carregando ferramenta...</div>';

    const resposta = await fetch(arquivo);

    if (!resposta.ok) {
      throw new Error(`Não foi possível carregar: ${arquivo}`);
    }

    const html = await resposta.text();
    app.innerHTML = html;
    pageTitle.textContent = titulo;

    menuButtons.forEach((botao) => botao.classList.remove("active"));

    if (botaoAtivo) {
      botaoAtivo.classList.add("active");
    }

        if (arquivo.includes("geradorCartao.html")) {
      iniciarGeradorCartao();
    }

    iniciarFerramentaAtual(arquivo);
  } catch (erro) {
    app.innerHTML = `
      <section class="pagina ativa">
        <div class="tool-card">
          <h1>Erro ao carregar ferramenta</h1>
          <p>${erro.message}</p>
          <p>Confira se o arquivo existe e se o projeto está aberto pelo Live Server.</p>
        </div>
      </section>
    `;
  }
}

function iniciarFerramentaAtual(arquivo) {
  if (arquivo.includes("geradorBD.html")) {
   // iniciarGeradorBD();
  }
}

menuButtons.forEach((botao) => {
  botao.addEventListener("click", () => {
    carregarFerramenta(botao.dataset.file, botao.dataset.title, botao);
  });
});

// Carregamento inicial
const primeiroBotao = document.querySelector(".menu-btn.active") || menuButtons[0];
carregarFerramenta(primeiroBotao.dataset.file, primeiroBotao.dataset.title, primeiroBotao);

/* =====================================================
   GERADOR DE BD - TOTALMENTE LOCAL
   Lê várias planilhas .xlsx, empilha os dados e acrescenta
   a coluna ARQUIVO_ORIGEM ao final.
===================================================== */
function iniciarGeradorBD() {
  const qtdQuestoes = document.getElementById("qtdQuestoes");
  const fileInput = document.getElementById("fileInput");
  const fileCounter = document.getElementById("fileCounter");
  const btnLimpar = document.getElementById("btnLimpar");
  const listaVisual = document.getElementById("listaVisual");
  const btnMain = document.getElementById("btnMain");
  const loader = document.getElementById("loader");
  const status = document.getElementById("status");

  if (!fileInput || !btnMain || !listaVisual) return;

  let arquivos = [];

  fileInput.addEventListener("change", () => {
    const novosArquivos = Array.from(fileInput.files || []).filter((arquivo) =>
      arquivo.name.toLowerCase().endsWith(".xlsx")
    );

    novosArquivos.forEach((arquivo) => {
      const jaExiste = arquivos.some(
        (item) => item.name === arquivo.name && item.size === arquivo.size
      );

      if (!jaExiste) {
        arquivos.push(arquivo);
      }
    });

    fileInput.value = "";
    atualizarFila();
  });

  btnLimpar.addEventListener("click", () => {
    arquivos = [];
    limparStatus();
    atualizarFila();
  });

  btnMain.addEventListener("click", async () => {
    await gerarBancoDeDados(arquivos, qtdQuestoes.value);
  });

  function atualizarFila() {
    const total = arquivos.length;

    fileCounter.textContent = `${total} ${total === 1 ? "arquivo" : "arquivos"}`;
    btnMain.disabled = total === 0;
    btnLimpar.hidden = total === 0;

    if (total === 0) {
      listaVisual.innerHTML = `
        <div class="empty-file-list">
          <span>🗁</span>
          <p>Nenhum arquivo selecionado.</p>
        </div>
      `;
      return;
    }

    listaVisual.innerHTML = arquivos
      .map(
        (arquivo, indice) => `
          <div class="file-item">
            <div>
              <strong>${escaparHtml(arquivo.name)}</strong>
              <small>${formatarTamanho(arquivo.size)}</small>
            </div>
            <button class="remove-file-btn" type="button" data-index="${indice}">×</button>
          </div>
        `
      )
      .join("");

    listaVisual.querySelectorAll(".remove-file-btn").forEach((botao) => {
      botao.addEventListener("click", () => {
        const indice = Number(botao.dataset.index);
        arquivos.splice(indice, 1);
        atualizarFila();
      });
    });
  }

  async function gerarBancoDeDados(listaArquivos, quantidadeQuestoes) {
    if (!window.XLSX) {
      mostrarStatus("A biblioteca xlsx.full.min.js não foi carregada.", "error");
      return;
    }

    if (listaArquivos.length === 0) {
      mostrarStatus("Selecione pelo menos uma planilha .xlsx.", "error");
      return;
    }

    if (loader) loader.hidden = false;
    btnMain.disabled = true;
    limparStatus();

    try {
      let cabecalhoBase = null;
      const linhasConsolidadas = [];
      const resumoArquivos = [];

      for (const arquivo of listaArquivos) {
        const dados = await lerArquivoComoArrayBuffer(arquivo);
        const workbook = XLSX.read(dados, { type: "array" });
        const primeiraAba = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[primeiraAba];

        const linhas = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: "",
          raw: false,
        });

        const linhasLimpas = removerLinhasVazias(linhas);

        if (linhasLimpas.length === 0) {
          resumoArquivos.push([arquivo.name, 0, "Arquivo ignorado: sem dados"]);
          continue;
        }

        const indiceCabecalho = encontrarCabecalho(linhasLimpas);
        const cabecalhoAtual = normalizarLinha(linhasLimpas[indiceCabecalho]);
        const dadosArquivo = linhasLimpas.slice(indiceCabecalho + 1);

        if (!cabecalhoBase) {
          cabecalhoBase = [...cabecalhoAtual, "ARQUIVO_ORIGEM"];
        }

        let linhasValidas = 0;

        dadosArquivo.forEach((linha) => {
          const linhaNormalizada = normalizarLinha(linha, cabecalhoBase.length - 1);

          if (linhaNormalizada.some((celula) => String(celula).trim() !== "")) {
            linhasConsolidadas.push([...linhaNormalizada, arquivo.name]);
            linhasValidas += 1;
          }
        });

        resumoArquivos.push([arquivo.name, linhasValidas, "Processado"]);
      }

      if (!cabecalhoBase || linhasConsolidadas.length === 0) {
        throw new Error("Nenhuma linha de dados foi encontrada nas planilhas selecionadas.");
      }

      const workbookSaida = XLSX.utils.book_new();
      const dadosSaida = [cabecalhoBase, ...linhasConsolidadas];
      const wsBD = XLSX.utils.aoa_to_sheet(dadosSaida);
      XLSX.utils.book_append_sheet(workbookSaida, wsBD, "BD");

      const info = [
        ["INFORMAÇÃO", "VALOR"],
        ["Quantidade de questões selecionada", quantidadeQuestoes],
        ["Arquivos processados", listaArquivos.length],
        ["Linhas consolidadas", linhasConsolidadas.length],
        ["Gerado em", new Date().toLocaleString("pt-BR")],
        [],
        ["ARQUIVO", "LINHAS", "STATUS"],
        ...resumoArquivos,
      ];

      const wsInfo = XLSX.utils.aoa_to_sheet(info);
      XLSX.utils.book_append_sheet(workbookSaida, wsInfo, "INFO");

      const nomeArquivo = `BD_GERADO_${quantidadeQuestoes}_QUESTOES.xlsx`;
      XLSX.writeFile(workbookSaida, nomeArquivo);

      mostrarStatus(
        `Banco de dados gerado com sucesso: ${linhasConsolidadas.length} linhas consolidadas.`,
        "success"
      );
    } catch (erro) {
      mostrarStatus(`Erro: ${erro.message}`, "error");
    } finally {
      if (loader) loader.hidden = true;
      btnMain.disabled = arquivos.length === 0;
    }
  }

  atualizarFila();

  function mostrarStatus(mensagem, tipo) {
    status.textContent = mensagem;
    status.className = `status-area ${tipo}`;
  }

  function limparStatus() {
    status.textContent = "";
    status.className = "status-area";
  }
}

function lerArquivoComoArrayBuffer(arquivo) {
  return new Promise((resolve, reject) => {
    const leitor = new FileReader();
    leitor.onload = (evento) => resolve(evento.target.result);
    leitor.onerror = () => reject(new Error(`Não foi possível ler ${arquivo.name}`));
    leitor.readAsArrayBuffer(arquivo);
  });
}

function removerLinhasVazias(linhas) {
  return linhas.filter((linha) =>
    linha.some((celula) => String(celula ?? "").trim() !== "")
  );
}

function encontrarCabecalho(linhas) {
  const indice = linhas.findIndex(
    (linha) => linha.filter((celula) => String(celula ?? "").trim() !== "").length >= 2
  );

  return indice >= 0 ? indice : 0;
}

function normalizarLinha(linha, tamanho = null) {
  const novaLinha = Array.from(linha || []).map((celula) =>
    typeof celula === "string" ? celula.trim() : celula
  );

  if (tamanho === null) {
    return novaLinha;
  }

  while (novaLinha.length < tamanho) {
    novaLinha.push("");
  }

  return novaLinha.slice(0, tamanho);
}

function formatarTamanho(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function escaparHtml(texto) {
  return String(texto)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/* script do Gerador de cartão-resposta*/

function iniciarGeradorCartao() {
  const pdfInput = document.getElementById("pdfModeloCartao");
  const xlsxInput = document.getElementById("xlsxCartaoInput");
  const listaVisual = document.getElementById("listaCartaoVisual");
  const fileCounter = document.getElementById("cartaoFileCounter");
  const btnLimpar = document.getElementById("btnLimparCartao");
  const btnGerar = document.getElementById("btnGerarCartoes");
  //const loader = document.getElementById("loaderCartao");
  const status = document.getElementById("statusCartao");

  if (!pdfInput || !xlsxInput || !listaVisual || !btnGerar) {
    console.warn("Elementos do Gerador de Cartão não encontrados.");
    return;
  }

  let pdfModelo = null;
  let filaPlanilhas = [];

  const mm = (v) => v * 2.83465;

  const COL_ESTUDANTE = 2;
  const LINHA_INICIAL_DADOS = 1;

  const CALIBRA = {
    ESTUDANTE: { x: mm(20), y: mm(167.2), sz: 12 },
    TURMA: { x: mm(20), y: mm(160.1), sz: 12 },
    ESCOLA: { x: mm(20), y: mm(153.0), sz: 12 },
  };


  const AREAS_LIMPEZA = {
    ESTUDANTE: { x: mm(28.5), y: mm(162.7), w: mm(175), h: mm(6.7) },
    TURMA: { x: mm(28.5), y: mm(155.6), w: mm(175), h: mm(6.7) },
    ESCOLA: { x: mm(28.5), y: mm(148.5), w: mm(175), h: mm(6.7) },
  };

  function atualizarBotao() {
    btnGerar.disabled = !(pdfModelo && filaPlanilhas.length > 0);
    fileCounter.textContent =
      filaPlanilhas.length === 1
        ? "1 arquivo"
        : `${filaPlanilhas.length} arquivos`;

    btnLimpar.hidden = filaPlanilhas.length === 0;
  }

  function renderLista() {
    if (filaPlanilhas.length === 0) {
      listaVisual.innerHTML = `
        <div class="empty-file-list">
          <span>🗁</span>
          <p>Nenhuma planilha selecionada.</p>
        </div>
      `;
      atualizarBotao();
      return;
    }

    listaVisual.innerHTML = filaPlanilhas
      .map((arquivo, indice) => {
        const info = extrairEscolaTurmaDoNome(arquivo.name);

        return `
          <div class="file-item">
            <div>
              <strong>📄 ${arquivo.name}</strong>
              <small>Escola: ${info.escola}</small>
              <small>Turma: ${info.turma}</small>
            </div>

            <button
              type="button"
              class="remove-file-btn"
              data-index="${indice}"
            >
              Remover
            </button>
          </div>
        `;
      })
      .join("");

    document.querySelectorAll(".remove-file-btn").forEach((botao) => {
      botao.addEventListener("click", () => {
        const indice = Number(botao.dataset.index);
        filaPlanilhas.splice(indice, 1);
        renderLista();
      });
    });

    atualizarBotao();
  }

  function limparNomeArquivo(nome) {
    return String(nome || "Cartoes")
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function valorLinha(linha, indice) {
    return String(linha[indice] || "").trim();
  }

  function extrairEscolaTurmaDoNome(nomeArquivo) {
    let nome = String(nomeArquivo || "")
      .replace(/\.xlsx$/i, "")
      .trim();

    nome = nome.replace(/^RELAÇÃO DE ALUNOS\s*-\s*/i, "").trim();

    const partes = nome.split(/\s+-\s+/);

    let escola = "";
    let turma = "";

    if (partes.length >= 2) {
      turma = partes.pop().trim();
      escola = partes.join(" - ").trim();
    } else {
      escola = "Escola não identificada";
      turma = nome.trim();
    }

    turma = turma
      .split("_")
      .map((parte) => parte.trim())
      .filter(Boolean)
      .map((parte) => {
        if (parte.toLowerCase() === "integra") return "Integral";
        return parte;
      })
      .join(" | ");

    return { escola, turma };
  }

  function limparCamposDoModelo(page, rgb) {
    Object.values(AREAS_LIMPEZA).forEach((area) => {
      page.drawRectangle({
        x: area.x,
        y: area.y,
        width: area.w,
        height: area.h,
        color: rgb(1, 1, 1),
      });
    });
  }

  pdfInput.addEventListener("change", () => {
    pdfModelo = pdfInput.files[0] || null;
    atualizarBotao();

    if (pdfModelo) {
      status.textContent = `PDF modelo selecionado: ${pdfModelo.name}`;
    } else {
      status.textContent = "";
    }
  });

  xlsxInput.addEventListener("change", () => {
    const novosArquivos = Array.from(xlsxInput.files || []);
    filaPlanilhas.push(...novosArquivos);
    xlsxInput.value = "";
    renderLista();
  });

  btnLimpar.addEventListener("click", () => {
    filaPlanilhas = [];
    renderLista();
    status.textContent = "";
  });

  btnGerar.addEventListener("click", async () => {
    try {
      if (!pdfModelo) {
        alert("Envie o PDF modelo do cartão.");
        return;
      }

      if (filaPlanilhas.length === 0) {
        alert("Envie pelo menos uma planilha .xlsx.");
        return;
      }

      if (!window.PDFLib) {
        throw new Error("A biblioteca pdf-lib não foi carregada.");
      }

      if (!window.JSZip) {
        throw new Error("A biblioteca JSZip não foi carregada.");
      }

     /* btnGerar.disabled = true;
      loader.hidden = false;
      status.textContent = "Lendo PDF modelo...";*/

      const { PDFDocument, rgb, StandardFonts } = PDFLib;
      const zip = new JSZip();

      const modeloBytes = await pdfModelo.arrayBuffer();

      for (const arquivo of filaPlanilhas) {
        status.textContent = `Lendo planilha: ${arquivo.name}`;

        const info = extrairEscolaTurmaDoNome(arquivo.name);

        const data = await arquivo.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
        });

        const pdfFinal = await PDFDocument.create();
        const pdfModeloBase = await PDFDocument.load(modeloBytes);
        const fonte = await pdfFinal.embedFont(StandardFonts.Helvetica);

        let totalGerado = 0;

        for (let i = LINHA_INICIAL_DADOS; i < rows.length; i++) {
          const linha = rows[i];
          const estudante = valorLinha(linha, COL_ESTUDANTE);

          if (!estudante) continue;

          const [pagina] = await pdfFinal.copyPages(pdfModeloBase, [0]);
          pdfFinal.addPage(pagina);

          //limparCamposDoModelo(pagina, rgb);

          pagina.drawText(`ESTUDANTE: ${estudante}`, {
            x: CALIBRA.ESTUDANTE.x,
            y: CALIBRA.ESTUDANTE.y,
            size: CALIBRA.ESTUDANTE.sz,
            font: fonte,
            color: rgb(0, 0, 0),
          });

          pagina.drawText(`TURMA: ${info.turma}`, {
            x: CALIBRA.TURMA.x,
            y: CALIBRA.TURMA.y,
            size: CALIBRA.TURMA.sz,
            font: fonte,
            color: rgb(0, 0, 0),
          });

          pagina.drawText(`ESCOLA: ${info.escola}`, {
            x: CALIBRA.ESCOLA.x,
            y: CALIBRA.ESCOLA.y,
            size: CALIBRA.ESCOLA.sz,
            font: fonte,
            color: rgb(0, 0, 0),
          });

          totalGerado++;
        }

        if (totalGerado > 0) {
          const pdfBytes = await pdfFinal.save();

          const nomePdf = limparNomeArquivo(
            `Cartoes - ${info.turma} - ${info.escola}`
          );

          zip.file(`${nomePdf}.pdf`, pdfBytes);
        }
      }

      status.textContent = "Gerando arquivo ZIP...";

      const zipBlob = await zip.generateAsync({ type: "blob" });

      const link = document.createElement("a");
      link.href = URL.createObjectURL(zipBlob);
      link.download = "Cartoes_Resposta.zip";
      link.click();

      status.textContent = "Concluído! ZIP baixado com sucesso.";

      filaPlanilhas = [];
      renderLista();
    } catch (erro) {
      console.error(erro);
      status.textContent = `Erro: ${erro.message}`;
    } finally {
      loader.hidden = true;
      atualizarBotao();
    }
  });

  renderLista();
}
