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
    iniciarGeradorBD();
  }

  if (arquivo.includes("uniPlan.html")) {
    iniciarUnificador();
  }
}


function iniciarGeradorBD() {
  const input = document.getElementById("fileInput");
  const btnMain = document.getElementById("btnMain");
  const btnLimpar = document.getElementById("btnLimpar");

  if (!input || !btnMain || !btnLimpar) return;

  input.addEventListener("change", adicionarArquivos);
  btnMain.addEventListener("click", gerarBD);
  btnLimpar.addEventListener("click", limparTudo);

  atualizarInterface();
}

function iniciarUnificador() {
  const input = document.getElementById("fileInput");
  const btnMain = document.getElementById("btnMain");
  const btnLimpar = document.getElementById("btnLimpar");

  if (!input || !btnMain || !btnLimpar) return;

  input.addEventListener("change", adicionarArquivosUnificador);
  btnMain.addEventListener("click", unificarPlanilhas);
  btnLimpar.addEventListener("click", limparTudoUnificador);

  atualizarInterfaceUnificador();
}

async function unificarPlanilhas() {
  if (arquivosUnificador.length === 0) {
    alert("Selecione pelo menos uma planilha.");
    return;
  }

  const dadosUnificados = [];
  let cabecalhoAdicionado = false;

  for (const arquivo of arquivosUnificador) {
    const buffer = await arquivo.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const primeiraAba = workbook.Sheets[workbook.SheetNames[0]];
    const linhas = XLSX.utils.sheet_to_json(primeiraAba, {
      header: 1,
      defval: ""
    });

    if (linhas.length === 0) continue;

    if (!cabecalhoAdicionado) {
      dadosUnificados.push(...linhas);
      cabecalhoAdicionado = true;
    } else {
      dadosUnificados.push(...linhas.slice(1));
    }
  }

  if (dadosUnificados.length === 0) {
    alert("Nenhum dado encontrado nas planilhas.");
    return;
  }

  const novaPlanilha = XLSX.utils.book_new();
  const novaAba = XLSX.utils.aoa_to_sheet(dadosUnificados);
  XLSX.utils.book_append_sheet(novaPlanilha, novaAba, "UNIFICADO");
  XLSX.writeFile(novaPlanilha, "PLANILHAS_UNIFICADAS.xlsx");
}

function adicionarArquivosUnificador() {
  const input = document.getElementById("fileInput");

  for (const arquivo of input.files) {
    if (!arquivosUnificador.some(item => item.name === arquivo.name)) {
      arquivosUnificador.push(arquivo);
    }
  }

  input.value = "";
  atualizarInterfaceUnificador();
}

function atualizarInterfaceUnificador() {
  const lista = document.getElementById("listaVisual");
  const contador = document.getElementById("fileCounter");
  const btnMain = document.getElementById("btnMain");
  const btnLimpar = document.getElementById("btnLimpar");

  if (!lista || !contador || !btnMain || !btnLimpar) return;

  contador.textContent = arquivosUnificador.length === 1
    ? "1 arquivo"
    : `${arquivosUnificador.length} arquivos`;

  if (arquivosUnificador.length === 0) {
    lista.innerHTML = `
      <div class="empty-file-list">
        <span>📂</span>
        <p>Nenhum arquivo selecionado.</p>
      </div>
    `;
    btnMain.disabled = true;
    btnLimpar.hidden = true;
    return;
  }

  lista.innerHTML = arquivosUnificador.map((arquivo, i) => `
    <div class="file-item">
      <span>📄 ${arquivo.name}</span>
      <button type="button" class="remove-file-btn" data-unificador-index="${i}">Remover</button>
    </div>
  `).join("");

  lista.querySelectorAll("[data-unificador-index]").forEach(botao => {
    botao.addEventListener("click", () => {
      removerArquivoUnificador(Number(botao.dataset.unificadorIndex));
    });
  });

  btnMain.disabled = false;
  btnLimpar.hidden = false;
}

function removerArquivoUnificador(i) {
  arquivosUnificador.splice(i, 1);
  atualizarInterfaceUnificador();
}

function limparTudoUnificador() {
  arquivosUnificador = [];
  atualizarInterfaceUnificador();
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
  let arquivosAcumulados = [];
  let arquivosUnificador = [];


    // 7 dados fixos do aluno + colunas das questoes
    const COLUMNS_D8 = [
      "A","E","C","D","H","I","F",
      "U","X","AA","AD","AG","AJ","AM","AP","AS","AV","AY",
      "BB","BE","BH","BK","BN","BQ","BT","BW","BZ",
      "CC","CF","CI","CL","CO","CR","CU","CX",
      "DA","DD","DG","DJ","DM","DP","DS","DV","DY",
      "EB","EE","EH","EK","EN","EQ","ET","EW","EZ",
      "FC","FF","FI","FL","FO","FR"
    ];

    // Mantendo a mesma lógica:
    // 22 questões => 29 colunas (7 fixas + 22)
    // 26 questões => 33 colunas (7 fixas + 26)
    // 44 questões => 51 colunas (7 fixas + 44)
    // 52 questões => 59 colunas (7 fixas + 52)
    const CONFIG_QUESTOES = {
  22: 29,
  26: 33,
  44: 51,
  52: 59
};
    function colToIdx(letter) {
      let column = 0;
      for (let i = 0; i < letter.length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
      }
      return column - 1;
    }

    function adicionarArquivos() {
      const input = document.getElementById('fileInput');
      for (let f of input.files) {
        if (!arquivosAcumulados.some(x => x.name === f.name)) {
          arquivosAcumulados.push(f);
        }
      }
      atualizarInterface();
      input.value = "";
    }

    function atualizarInterface() {
      const lista = document.getElementById('listaVisual');
      const btnG = document.getElementById('btnMain');
      const btnL = document.getElementById('btnLimpar');

      if (arquivosAcumulados.length === 0) {
        lista.innerHTML = '<div class="text-center text-muted py-3">Vazio</div>';
        btnG.disabled = true;
        btnL.style.display = "none";
        return;
      }

      btnG.disabled = false;
      btnL.style.display = "block";
      lista.innerHTML = arquivosAcumulados.map((f, i) => `
        <div class="file-item">
          <span>📄 ${f.name}</span>
          <span class="remove-btn" onclick="removerArquivo(${i})">✕</span>
        </div>
      `).join('');
    }

    function removerArquivo(i) {
      arquivosAcumulados.splice(i, 1);
      atualizarInterface();
    }

    function limparTudo() {
      arquivosAcumulados = [];
      atualizarInterface();
    }

function normalizarResposta(valor) {
  if (valor === undefined || valor === null || String(valor).trim() === "") {
    return "∅";
  }

  const texto = String(valor).toUpperCase().trim();

  if (texto === "∅") {
    return "∅";
  }

  if (texto === "#") {
    return "#";
  }

  // Captura alternativas A, B, C, D ou E encontradas no campo
  const marcadas = texto.match(/[A-E]/g) || [];

  // Remove repetidas, caso apareça algo como "A A"
  const unicas = [...new Set(marcadas)];

  if (unicas.length === 0) {
    return "∅";
  }

  if (unicas.length === 1) {
    return unicas[0];
  }

  // Se marcou mais de uma alternativa
  return "#";
}

    async function gerarBD() {
      const loader = document.getElementById('loader');
      const btnG = document.getElementById('btnMain');
      const qtd = parseInt(document.getElementById('qtdQuestoes').value, 10);

      if (loader) loader.style.display = "block";
      btnG.disabled = true;

      const totalColunas = CONFIG_QUESTOES[qtd];

      if (!totalColunas) {
        alert("Quantidade de questões inválida.");
        if (loader) loader.style.display = "none";
        btnG.disabled = false;
        return;
      }

      const colIndices = COLUMNS_D8.slice(0, totalColunas).map(colToIdx);
      let dadosFinais = [];

      try {
        for (let file of arquivosAcumulados) {
          const data = await file.arrayBuffer();
          const workbook = XLSX.read(data);
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

          let linesInFile = 0;

          for (let i = 1; i < rows.length; i++) {
            let row = rows[i];

            if (row && row[5] && String(row[5]).trim() !== "") {
              let extracted = colIndices.map(idx =>
                (row[idx] !== undefined && row[idx] !== "") ? row[idx] : "∅"
              );

              let newRow = new Array(totalColunas).fill("");

              // 7 primeiras colunas: dados do aluno na ordem A, E, C, D, H, I, F
              for (let j = 0; j < 7; j++) {
                 newRow[j] = extracted[j];
                 }

              for (let j = 7; j < extracted.length; j++) {
  newRow[j] = normalizarResposta(extracted[j]);
}

              dadosFinais.push(newRow);
              linesInFile++;
            }
          }

          if (linesInFile > 0) {
            let sep = new Array(totalColunas).fill("");
            sep[0] = `--- FIM DO ARQUIVO: ${file.name} ---`;
            dadosFinais.push(sep);
          }
        }

if (dadosFinais.length === 0) {
  alert("Nenhum aluno válido foi encontrado nas planilhas selecionadas.");
  btnG.disabled = false;
  return;
}

        // Geração 100% local do banco de dados (.xlsx)
        const wbFinal = XLSX.utils.book_new();
        const wsFinal = XLSX.utils.aoa_to_sheet(dadosFinais);
        XLSX.utils.book_append_sheet(wbFinal, wsFinal, "RESPOSTAS");

        const agora = new Date();
        const pad = n => String(n).padStart(2, "0");
        const nomeArquivo = `BD_ULTRA_RAPIDO_${pad(agora.getDate())}-${pad(agora.getMonth() + 1)}-${agora.getFullYear()}_${pad(agora.getHours())}${pad(agora.getMinutes())}.xlsx`;

        XLSX.writeFile(wbFinal, nomeArquivo);

        if (loader) loader.style.display = "none";
        const status = document.getElementById("status");
        if (status) {
          status.innerHTML = `<div class="alert alert-success mt-3"><b>Sucesso!</b><br>Banco de dados baixado: ${nomeArquivo}</div>`;
        }

        arquivosAcumulados = [];
        atualizarInterface();

      } catch (e) {
        alert("Erro ao ler arquivos: " + e.message);
        if (loader) loader.style.display = "none";
        btnG.disabled = false;
      }
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
          <span>📂</span>
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
      atualizarBotao();
    }
  });

  renderLista();
}
