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
    // geradorBD();
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
  let arquivosAcumulados = [];

    // Mapeamento original D8 do professor
    const COLUMNS_D8 = [
      "A","C","D",
      "S","V","Y","AB","AE","AH","AK","AN","AQ","AT","AW","AZ","BC","BF","BI","BL","BO","BR","BU","BX","CA","CD","CG","CJ","CM","CP","CS","CV","CY","DB","DE","DH","DK","DN","DQ","DT","DW","DZ","EC","EF","EI","EL","EO","ER","EU","EX","FA","FD","FG","FJ","FM","FP"
    ];

    // Mantendo a mesma lógica:
    // 22 questões => 25 colunas (3 fixas + 22)
    // 26 questões => 29 colunas (3 fixas + 26)
    // 44 questões => 47 colunas (3 fixas + 44)
    // 52 questões => 55 colunas (3 fixas + 52)
    const CONFIG_QUESTOES = {
      22: 25,
      26: 29,
      44: 47,
      52: 55
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

      loader.style.display = "block";
      btnG.disabled = true;

      const totalColunas = CONFIG_QUESTOES[qtd];

      if (!totalColunas) {
        alert("Quantidade de questões inválida.");
        loader.style.display = "none";
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

          for (let i = 2; i < rows.length; i++) {
            let row = rows[i];

            if (row && row[2] && String(row[2]).trim() !== "") {
              let extracted = colIndices.map(idx =>
                (row[idx] !== undefined && row[idx] !== "") ? row[idx] : "∅"
              );

              let newRow = new Array(totalColunas + 3).fill("");

              // Mantida a sua lógica:
              // C -> A
              // A -> C
              // D -> F
              newRow[0] = extracted[1];
              newRow[2] = extracted[0];
              newRow[5] = extracted[2];

              for (let j = 3; j < extracted.length; j++) {
  newRow[j + 3] = normalizarResposta(extracted[j]);
}

              dadosFinais.push(newRow);
              linesInFile++;
            }
          }

          if (linesInFile > 0) {
            let sep = new Array(totalColunas + 3).fill("");
            sep[0] = `--- FIM DO ARQUIVO: ${file.name} ---`;
            dadosFinais.push(sep);
          }
        }

        google.script.run
          .withSuccessHandler(url => {
            loader.style.display = "none";
            document.getElementById('status').innerHTML =
              `<div class="alert alert-success mt-3"><b>Sucesso!</b><br><a href="${url}" target="_blank" class="btn btn-success mt-2">ABRIR RESULTADO</a></div>`;
            arquivosAcumulados = [];
            atualizarInterface();
          })
          .withFailureHandler(err => {
            alert(err);
            loader.style.display = "none";
            btnG.disabled = false;
          })
          .criarPlanilhaFinal(dadosFinais);

      } catch (e) {
        alert("Erro ao ler arquivos: " + e.message);
        loader.style.display = "none";
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
      loader.hidden = true;
      atualizarBotao();
    }
  });

  renderLista();
}
