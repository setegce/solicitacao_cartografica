(() => {
  "use strict";

  // ================= CONFIG =================
  const CODIGO_GESTOR = "061098";
  const CODIGOS_TECNICOS = {
    LAIS: "061098",
    LAIZE: "061098",
    VALESKA: "061098",
    LIZABETH: "061098",
    ISMAEL: "061098",
    FERNANDO: "061098",
  };

  // ================= STATE =================
  let acessoGestor = false;
  let acessoTecnico = false;
  let tecnicoLogado = null;

  let solicitacoes = [];
  let firebaseReady = false;

  // ================= DOM =================
  let formSection,
    solicitacaoForm,
    tbodySolicitacoes,
    emptyState,
    tabelaSolicitacoes;

  let modalDetalhes,
    conteudoDetalhes,
    modalAcessoGestor,
    modalAcessoTecnico,
    modalAtribuicao,
    conteudoAtribuicao,
    modalRelatorio,
    modalConfirmacao,
    conteudoConfirmacao;

  let painelGestor,
    painelTecnico,
    tituloPainelTecnico,
    listaTecnicos,
    estatisticasTecnico,
    cardNovaSolicitacao,
    headerButtons;

  let tabBtns;

  // ================= INIT =================
  document.addEventListener("DOMContentLoaded", async () => {
    console.log("🚀 Sistema - Iniciando...");

    // DOM refs
    formSection = document.getElementById("formSection");
    solicitacaoForm = document.getElementById("solicitacaoForm");
    tbodySolicitacoes = document.getElementById("tbodySolicitacoes");
    emptyState = document.getElementById("emptyState");
    tabelaSolicitacoes = document.getElementById("tabelaSolicitacoes");

    modalDetalhes = document.getElementById("modalDetalhes");
    conteudoDetalhes = document.getElementById("conteudoDetalhes");

    modalAcessoGestor = document.getElementById("modalAcessoGestor");
    modalAcessoTecnico = document.getElementById("modalAcessoTecnico");

    modalAtribuicao = document.getElementById("modalAtribuicao");
    conteudoAtribuicao = document.getElementById("conteudoAtribuicao");

    modalRelatorio = document.getElementById("modalRelatorio");

    modalConfirmacao = document.getElementById("modalConfirmacao");
    conteudoConfirmacao = document.getElementById("conteudoConfirmacao");

    painelGestor = document.getElementById("painelGestor");
    painelTecnico = document.getElementById("painelTecnico");
    tituloPainelTecnico = document.getElementById("tituloPainelTecnico");

    listaTecnicos = document.getElementById("listaTecnicos");
    estatisticasTecnico = document.getElementById("estatisticasTecnico");

    cardNovaSolicitacao = document.getElementById("cardNovaSolicitacao");
    headerButtons = document.getElementById("headerButtons");

    tabBtns = document.querySelectorAll(".tab-btn");

    // Binds extras (se você colocou ids no HTML novo)
    document.getElementById("btnAcessoGestor")?.addEventListener("click", abrirModalAcessoGestor);
    document.getElementById("btnAcessoTecnico")?.addEventListener("click", abrirModalAcessoTecnico);
    document.getElementById("btnToggleForm")?.addEventListener("click", toggleForm);
    document.getElementById("btnLimparForm")?.addEventListener("click", limparForm);

    // Toggles do formulário
    document.getElementById("tipoMapa")?.addEventListener("change", toggleTipoMapaOutros);
    document.getElementById("elementoOutros")?.addEventListener("change", toggleElementosOutros);
    document.querySelectorAll('input[name="artNecessaria"]').forEach((el) => {
      el.addEventListener("change", toggleARTResponsavel);
    });

    // ENTER pra logar nos modais
    bindEnterNoModal(modalAcessoGestor, validarCodigoAcesso);
    bindEnterNoModal(modalAcessoTecnico, validarAcessoTecnico);

    // Enter nos inputs de login
    document.getElementById("codigoAcesso")?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        validarCodigoAcesso();
      }
    });

    document.getElementById("codigoAcessoTecnico")?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        validarAcessoTecnico();
      }
    });

    document.getElementById("selectTecnicoAcesso")?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        validarAcessoTecnico();
      }
    });

    // Datas padrão
    const hoje = new Date();
    setInputDate("dataSolicitacao", hoje);

    const entrega = new Date();
    entrega.setDate(hoje.getDate() + 15);
    setInputDate("dataEntrega", entrega);

    const inicioRel = new Date();
    inicioRel.setMonth(inicioRel.getMonth() - 1);
    setInputDate("relatorioPeriodoInicio", inicioRel);
    setInputDate("relatorioPeriodoFim", hoje);

    // Submit
    solicitacaoForm?.addEventListener("submit", async (e) => {
      e.preventDefault();
      if (!validarFormulario()) return;
      await salvarNovaSolicitacao();
    });

    // Tabs
    tabBtns?.forEach((btn) => {
      btn.addEventListener("click", function () {
        tabBtns.forEach((b) => b.classList.remove("active"));
        this.classList.add("active");

        const tab = this.dataset.tab;
        const map = {
          todas: null,
          fila: "fila",
          processando: "processando",
          aguardando: "aguardando",
          finalizado: "finalizado",
          finalizadas: "finalizado",
        };

        atualizarTabela(map[tab] ?? null);
      });
    });

    // Clique fora fecha
    window.addEventListener("click", (e) => {
      if (e.target === modalDetalhes) fecharModal();
      if (e.target === modalAtribuicao) fecharModalAtribuicao();
      if (e.target === modalConfirmacao) fecharModalConfirmacao();
      if (e.target === modalAcessoGestor) fecharModalAcessoGestor();
      if (e.target === modalAcessoTecnico) fecharModalAcessoTecnico();
      if (e.target === modalRelatorio) fecharModalRelatorio();
    });

    // ESC fecha
    window.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      fecharModal();
      fecharModalAtribuicao();
      fecharModalConfirmacao();
      fecharModalAcessoGestor();
      fecharModalAcessoTecnico();
      fecharModalRelatorio();
    });

    // Firebase
    await aguardarFirebase();

    atualizarTabela();
    atualizarVisibilidadeFormulario();

    console.log("✅ Sistema carregado!");
  });

  // ================= FIREBASE =================
  function isFirebaseOk() {
    return !!(window.firebaseApp && window.db && window.dbRef && window.firebaseFunctions);
  }

  function aguardarFirebase() {
    return new Promise((resolve) => {
      if (isFirebaseOk()) {
        carregarDados();
        resolve();
        return;
      }

      let tentativas = 0;
      const interval = setInterval(() => {
        tentativas++;
        if (isFirebaseOk()) {
          clearInterval(interval);
          carregarDados();
          resolve();
        } else if (tentativas > 40) {
          // ~20s
          console.error("❌ Firebase não inicializou em ~20s. Verifique se index.html está com scripts type=module.");
          mostrarNotificacao("❌ Firebase não inicializou. Verifique o console.", "error");
          clearInterval(interval);
          resolve();
        }
      }, 500);
    });
  }

  function carregarDados() {
    if (!window.db || !window.dbRef || !window.firebaseFunctions?.onValue) return;

    window.firebaseFunctions.onValue(
      window.dbRef,
      (snapshot) => {
        try {
          const data = snapshot.val();
          if (!data) {
            solicitacoes = [];
            firebaseReady = true;
            atualizarTabela();
            atualizarListaTecnicos();
            atualizarEstatisticas();
            atualizarEstatisticasTecnico();
            return;
          }

          // data = { "1": {...}, "2": {...} }
          solicitacoes = Object.entries(data)
            .map(([key, value]) => ({
              id: Number(key),
              ...value,
            }))
            .filter((s) => Number.isFinite(s.id))
            .sort((a, b) => (b.id ?? 0) - (a.id ?? 0));

          firebaseReady = true;

          atualizarTabela();
          atualizarListaTecnicos();
          atualizarEstatisticas();
          atualizarEstatisticasTecnico();
        } catch (err) {
          console.error("❌ Erro lendo dados:", err);
          mostrarNotificacao("Erro ao carregar dados.", "error");
        }
      },
      (error) => {
        console.error("❌ Firebase:", error);
        mostrarNotificacao("Falha na conexão com servidor.", "error");
      }
    );
  }

  // ID atômico: counters/solicitacoesNextId
  async function obterProximoIdAtomico() {
    const fn = window.firebaseFunctions;
    if (!fn?.ref || !fn?.runTransaction || !window.db) {
      // fallback (não ideal)
      const max = solicitacoes.length ? Math.max(...solicitacoes.map((s) => s.id)) : 0;
      return max + 1;
    }

    const counterRef = fn.ref(window.db, "counters/solicitacoesNextId");
    const result = await fn.runTransaction(counterRef, (current) => {
      const n = Number(current);
      if (!Number.isFinite(n) || n < 1) return 1;
      return n + 1;
    });

    return Number(result.snapshot.val());
  }

  function getSolicitacaoRef(id) {
    if (!window.db || !window.firebaseFunctions?.ref) return null;
    return window.firebaseFunctions.ref(window.db, `solicitacoes/${id}`);
  }

  function atualizarSolicitacaoFirebase(id, patch) {
    if (!firebaseReady) {
      mostrarNotificacao("⚠️ Aguardando conexão com servidor...", "warning");
      return;
    }

    const r = getSolicitacaoRef(id);
    if (!r || !window.firebaseFunctions?.update) {
      console.error("❌ Firebase update indisponível. Verifique firebase-config.js");
      mostrarNotificacao("❌ Erro de configuração do Firebase.", "error");
      return;
    }

    const payload = {
      ...patch,
      dataAtualizacao: new Date().toISOString(),
    };

    window.firebaseFunctions.update(r, payload).catch((err) => {
      console.error("❌ update:", err);
      mostrarNotificacao("❌ Falha ao atualizar solicitação.", "error");
    });
  }

  function excluirSolicitacao(id) {
    if (!acessoGestor) {
      mostrarNotificacao("⚠️ Apenas gestor pode excluir.", "warning");
      return;
    }

    const r = getSolicitacaoRef(id);
    if (!r || !window.firebaseFunctions?.remove) {
      console.error("❌ Firebase remove indisponível. Verifique firebase-config.js");
      mostrarNotificacao("❌ Erro de configuração do Firebase.", "error");
      return;
    }

    window.firebaseFunctions
      .remove(r)
      .then(() => {
        fecharModalConfirmacao();
        fecharModal();
        mostrarNotificacao(`🗑️ Solicitação #${id} excluída.`, "success");
      })
      .catch((err) => {
        console.error("❌ remove:", err);
        mostrarNotificacao("❌ Falha ao excluir solicitação.", "error");
      });
  }

  // ================= FORM / VALIDAÇÃO =================
  function toggleForm() {
    if (!formSection) return;
    formSection.classList.toggle("active");
    formSection.setAttribute("aria-hidden", String(!formSection.classList.contains("active")));
    if (formSection.classList.contains("active")) {
      formSection.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function limparForm() {
    solicitacaoForm?.reset();

    const hoje = new Date();
    setInputDate("dataSolicitacao", hoje);

    const entrega = new Date();
    entrega.setDate(hoje.getDate() + 15);
    setInputDate("dataEntrega", entrega);

    safeSetDisplay("tipoMapaOutros", "none");
    safeSetDisplay("artResponsavelContainer", "none");
    safeSetDisplay("elementosOutrosContainer", "none");
  }

  function toggleTipoMapaOutros() {
    const tipoMapa = document.getElementById("tipoMapa")?.value || "";
    const container = document.getElementById("tipoMapaOutros");
    if (!container) return;

    if (tipoMapa === "OUTROS") {
      container.style.display = "block";
    } else {
      container.style.display = "none";
      const input = document.getElementById("tipoMapaOutrosTexto");
      if (input) input.value = "";
    }
  }

  function toggleARTResponsavel() {
    const artNecessaria = document.querySelector('input[name="artNecessaria"]:checked')?.value;
    const container = document.getElementById("artResponsavelContainer");
    if (!container) return;

    if (artNecessaria === "sim") {
      container.style.display = "block";
    } else {
      container.style.display = "none";
      const input = document.getElementById("artResponsavel");
      if (input) input.value = "";
    }
  }

  function toggleElementosOutros() {
    const outrosCheckbox = document.getElementById("elementoOutros");
    const container = document.getElementById("elementosOutrosContainer");
    if (!outrosCheckbox || !container) return;

    if (outrosCheckbox.checked) {
      container.style.display = "block";
    } else {
      container.style.display = "none";
      const input = document.getElementById("elementosOutrosTexto");
      if (input) input.value = "";
    }
  }

  function validarFormulario() {
    const artNecessaria = document.querySelector('input[name="artNecessaria"]:checked')?.value;
    if (!artNecessaria) {
      mostrarNotificacao("⚠️ Selecione se é necessário solicitar ART!", "warning");
      return false;
    }

    const tipoMapa = document.getElementById("tipoMapa")?.value || "";
    if (!tipoMapa) {
      mostrarNotificacao("⚠️ Selecione o tipo de mapa!", "warning");
      document.getElementById("tipoMapa")?.focus();
      return false;
    }

    if (tipoMapa === "OUTROS") {
      const nomeTipoMapa = document.getElementById("tipoMapaOutrosTexto")?.value.trim() || "";
      if (!nomeTipoMapa) {
        mostrarNotificacao("⚠️ Digite o nome do tipo de mapa!", "warning");
        document.getElementById("tipoMapaOutrosTexto")?.focus();
        return false;
      }
    }

    const elementoOutros = document.getElementById("elementoOutros")?.checked;
    if (elementoOutros) {
      const txt = document.getElementById("elementosOutrosTexto")?.value.trim() || "";
      if (!txt) {
        mostrarNotificacao("⚠️ Descreva os elementos personalizados!", "warning");
        document.getElementById("elementosOutrosTexto")?.focus();
        return false;
      }
    }

    const obrigatorios = [
      { id: "solicitante", msg: "Solicitante" },
      { id: "cliente", msg: "Cliente" },
      { id: "nomeEstudo", msg: "Nome do Estudo" },
      { id: "localidade", msg: "Localidade" },
      { id: "dataEntrega", msg: "Data de Entrega" },
      { id: "diretorioArquivos", msg: "Diretório dos arquivos" },
      { id: "diretorioSalvamento", msg: "Diretório de salvamento" },
      { id: "finalidade", msg: "Finalidade" },
    ];

    for (const c of obrigatorios) {
      const el = document.getElementById(c.id);
      const val = el?.value?.trim() || "";
      if (!val) {
        mostrarNotificacao(`⚠️ Campo obrigatório: ${c.msg}`, "warning");
        el?.focus();
        return false;
      }
    }

    if (artNecessaria === "sim") {
      const resp = document.getElementById("artResponsavel")?.value.trim() || "";
      if (!resp) {
        mostrarNotificacao("⚠️ Informe o responsável técnico (ART).", "warning");
        document.getElementById("artResponsavel")?.focus();
        return false;
      }
    }

    return true;
  }

  // ================= SALVAR =================
  async function salvarNovaSolicitacao() {
    if (!firebaseReady) {
      mostrarNotificacao("⚠️ Aguardando conexão com servidor...", "warning");
      return false;
    }
    if (!window.db || !window.firebaseFunctions?.ref || !window.firebaseFunctions?.set) {
      mostrarNotificacao("❌ Firebase não configurado corretamente.", "error");
      return false;
    }

    const tipoMapa = document.getElementById("tipoMapa")?.value || "";
    const nomeTipoMapa = tipoMapa === "OUTROS" ? (document.getElementById("tipoMapaOutrosTexto")?.value.trim() || "") : "";

    const prazoDias = Number(document.getElementById("prazoDias")?.value) || 15;
    const dataSolic = document.getElementById("dataSolicitacao")?.value;

    const artNecessaria = document.querySelector('input[name="artNecessaria"]:checked')?.value || "nao";

    const novaSolicitacao = {
      solicitante: document.getElementById("solicitante")?.value.trim() || "",
      cliente: document.getElementById("cliente")?.value.trim() || "",
      empreendimento: document.getElementById("empreendimento")?.value.trim() || "Não informado",
      nomeEstudo: document.getElementById("nomeEstudo")?.value.trim() || "",
      localidade: document.getElementById("localidade")?.value.trim() || "",
      municipio: document.getElementById("municipio")?.value.trim() || "Não informado",
      dataSolicitacao: dataSolic,
      dataEntrega: document.getElementById("dataEntrega")?.value || "",
      tipoMapa,
      nomeTipoMapa,
      diretorioArquivos: document.getElementById("diretorioArquivos")?.value.trim() || "",
      diretorioSalvamento: document.getElementById("diretorioSalvamento")?.value.trim() || "",
      artNecessaria,
      artResponsavel: artNecessaria === "sim" ? (document.getElementById("artResponsavel")?.value.trim() || "") : "",
      finalidade: document.getElementById("finalidade")?.value || "",
      observacoes: document.getElementById("observacoes")?.value.trim() || "Nenhuma observação",
      prazoDias,
      status: "fila",
      tecnicoResponsavel: "PENDENTE",
      dataCriacao: new Date().toISOString(),
      dataAtualizacao: new Date().toISOString(),
      dataConclusaoPrevista: calcularDataConclusao(dataSolic, prazoDias),
      dataConclusaoReal: null,
      produtos: {
        mapa: Boolean(document.getElementById("produtoMapa")?.checked),
        croqui: Boolean(document.getElementById("produtoCroqui")?.checked),
        shapefile: Boolean(document.getElementById("produtoShapefile")?.checked),
        kml: Boolean(document.getElementById("produtoKML")?.checked),
      },
      elementos: {
        localizacao: Boolean(document.getElementById("elementoLocalizacao")?.checked),
        acessoLocal: Boolean(document.getElementById("elementoAcessoLocal")?.checked),
        acessoRegional: Boolean(document.getElementById("elementoAcessoRegional")?.checked),
        areaAmostral: Boolean(document.getElementById("elementoAreaAmostral")?.checked),
        outros: Boolean(document.getElementById("elementoOutros")?.checked),
        outrosTexto: document.getElementById("elementosOutrosTexto")?.value.trim() || "",
      },
    };

    try {
      const id = await obterProximoIdAtomico();
      const newRef = window.firebaseFunctions.ref(window.db, `solicitacoes/${id}`);
      await window.firebaseFunctions.set(newRef, novaSolicitacao);

      mostrarNotificacao("✅ Solicitação criada com sucesso!", "success");

      solicitacaoForm?.reset();
      formSection?.classList.remove("active");
      formSection?.setAttribute("aria-hidden", "true");

      safeSetDisplay("tipoMapaOutros", "none");
      safeSetDisplay("artResponsavelContainer", "none");
      safeSetDisplay("elementosOutrosContainer", "none");

      const hoje = new Date();
      setInputDate("dataSolicitacao", hoje);
      const entrega = new Date();
      entrega.setDate(hoje.getDate() + 15);
      setInputDate("dataEntrega", entrega);

      return true;
    } catch (err) {
      console.error("❌ salvarNovaSolicitacao:", err);
      mostrarNotificacao("❌ Erro ao salvar solicitação!", "error");
      return false;
    }
  }

  // ================= ACESSO =================
  function abrirModalAcessoGestor() {
    modalAcessoGestor?.classList.add("active");
    const inp = document.getElementById("codigoAcesso");
    if (inp) {
      inp.value = "";
      inp.focus();
    }
  }

  function fecharModalAcessoGestor() {
    modalAcessoGestor?.classList.remove("active");
  }

  function validarCodigoAcesso() {
    const codigo = document.getElementById("codigoAcesso")?.value || "";
    if (codigo === CODIGO_GESTOR) {
      acessoGestor = true;
      acessoTecnico = false;
      tecnicoLogado = null;

      fecharModalAcessoGestor();
      mostrarNotificacao("✅ Acesso de gestor concedido!", "success");

      if (painelGestor) painelGestor.style.display = "block";
      if (painelTecnico) painelTecnico.style.display = "none";

      if (headerButtons) {
        headerButtons.innerHTML = `
          <span class="chip">👤 Gestor</span>
          <button class="btn btn-ghost" type="button" onclick="fazerLogout()">Sair</button>
        `;
      }

      atualizarTabela();
      atualizarListaTecnicos();
      atualizarVisibilidadeFormulario();
    } else {
      mostrarNotificacao("❌ Código incorreto!", "error");
    }
  }

  function abrirModalAcessoTecnico() {
    modalAcessoTecnico?.classList.add("active");
    const sel = document.getElementById("selectTecnicoAcesso");
    const inp = document.getElementById("codigoAcessoTecnico");
    if (sel) sel.value = "";
    if (inp) inp.value = "";
    sel?.focus();
  }

  function fecharModalAcessoTecnico() {
    modalAcessoTecnico?.classList.remove("active");
  }

  function validarAcessoTecnico() {
    const tecnicoSelecionado = document.getElementById("selectTecnicoAcesso")?.value || "";
    const codigo = document.getElementById("codigoAcessoTecnico")?.value || "";

    if (!tecnicoSelecionado) {
      mostrarNotificacao("⚠️ Selecione seu nome!", "warning");
      return;
    }
    if (!codigo) {
      mostrarNotificacao("⚠️ Digite o código!", "warning");
      return;
    }

    if (codigo === CODIGOS_TECNICOS[tecnicoSelecionado]) {
      acessoTecnico = true;
      acessoGestor = false;
      tecnicoLogado = tecnicoSelecionado;

      fecharModalAcessoTecnico();
      mostrarNotificacao(`✅ Bem-vindo, ${formatarNomeTecnico(tecnicoLogado)}!`, "success");

      if (painelTecnico) painelTecnico.style.display = "block";
      if (painelGestor) painelGestor.style.display = "none";
      if (tituloPainelTecnico) {
        tituloPainelTecnico.textContent = `👨‍💼 Painel de ${formatarNomeTecnico(tecnicoLogado)}`;
      }

      if (headerButtons) {
        headerButtons.innerHTML = `
          <span class="chip">👤 ${formatarNomeTecnico(tecnicoLogado)}</span>
          <button class="btn btn-ghost" type="button" onclick="fazerLogout()">Sair</button>
        `;
      }

      atualizarEstatisticasTecnico();
      atualizarTabela();
      atualizarVisibilidadeFormulario();
    } else {
      mostrarNotificacao("❌ Código incorreto!", "error");
    }
  }

  function fazerLogout() {
    acessoGestor = false;
    acessoTecnico = false;
    tecnicoLogado = null;

    if (painelGestor) painelGestor.style.display = "none";
    if (painelTecnico) painelTecnico.style.display = "none";

    if (headerButtons) {
      headerButtons.innerHTML = `
        <button class="btn btn-ghost" type="button" id="btnAcessoGestor">🔐 Acesso Gestor</button>
        <button class="btn btn-primary" type="button" id="btnAcessoTecnico">👤 Acesso Técnico</button>
      `;
      headerButtons.querySelector("#btnAcessoGestor")?.addEventListener("click", abrirModalAcessoGestor);
      headerButtons.querySelector("#btnAcessoTecnico")?.addEventListener("click", abrirModalAcessoTecnico);
    }

    atualizarTabela();
    atualizarVisibilidadeFormulario();
    mostrarNotificacao("✅ Logout realizado!", "info");
  }

  function atualizarVisibilidadeFormulario() {
    if (!cardNovaSolicitacao) return;
    cardNovaSolicitacao.style.display = acessoGestor || acessoTecnico ? "none" : "block";
  }

  // ================= TABELA =================
  function atualizarTabela(filtroStatus = null) {
    if (!tbodySolicitacoes) return;

    tbodySolicitacoes.innerHTML = "";

    let dadosFiltrados = [...solicitacoes];

    if (acessoTecnico && tecnicoLogado) {
      dadosFiltrados = dadosFiltrados.filter((s) => s.tecnicoResponsavel === tecnicoLogado);
    }

    if (typeof filtroStatus === "string") {
      const map = {
        todas: null,
        fila: "fila",
        processando: "processando",
        aguardando: "aguardando",
        finalizado: "finalizado",
        finalizadas: "finalizado",
      };
      filtroStatus = map[filtroStatus] ?? filtroStatus;
    }

    if (filtroStatus) {
      dadosFiltrados = dadosFiltrados.filter((s) => s.status === filtroStatus);
    }

    if (dadosFiltrados.length === 0) {
      emptyState?.classList.add("active");
      if (tabelaSolicitacoes) tabelaSolicitacoes.style.display = "none";
    } else {
      emptyState?.classList.remove("active");
      if (tabelaSolicitacoes) tabelaSolicitacoes.style.display = "table";

      dadosFiltrados.sort((a, b) => (b.id ?? 0) - (a.id ?? 0));

      for (const s of dadosFiltrados) {
        const row = document.createElement("tr");

        const dataFormatada = formatDateBR(s.dataSolicitacao);
        const nomeTecnico = formatarNomeTecnico(s.tecnicoResponsavel);
        const municipioFormatado =
          s.municipio && s.municipio !== "Não informado" ? String(s.municipio).split("/")[0] : "Não informado";

        let botoesAcoes = `
          <button class="btn action-btn btn-info" type="button" onclick="verDetalhes(${s.id})">👁️ Ver</button>
        `;

        if (acessoGestor) {
          botoesAcoes += `
            <button class="btn action-btn btn-warning" type="button" onclick="abrirModalAtribuicao(${s.id})">👤 Atribuir</button>
            <button class="btn action-btn btn-success" type="button" onclick="mudarStatus(${s.id}, 'processando')">⚙️ Processar</button>
          `;
        } else if (acessoTecnico && s.tecnicoResponsavel === tecnicoLogado) {
          botoesAcoes += `
            <button class="btn action-btn btn-warning" type="button" onclick="mudarStatus(${s.id}, 'processando')">⚙️ Processar</button>
            <button class="btn action-btn btn-info" type="button" onclick="mudarStatus(${s.id}, 'aguardando')">⏳ Aguardar</button>
            <button class="btn action-btn btn-success" type="button" onclick="finalizarSolicitacao(${s.id})">✅ Finalizar</button>
          `;
        }

        row.innerHTML = `
          <td><strong>${String(s.id).padStart(4, "0")}</strong></td>
          <td>${escapeHtml(s.solicitante || "")}</td>
          <td>${escapeHtml(s.nomeEstudo || "Não informado")}</td>
          <td>${escapeHtml(s.cliente || "")}</td>
          <td>${escapeHtml(formatarTipoMapa(s.tipoMapa, s.nomeTipoMapa))}</td>
          <td>${escapeHtml(municipioFormatado)}</td>
          <td>${
            acessoGestor
              ? escapeHtml(nomeTecnico)
              : s.tecnicoResponsavel !== "PENDENTE"
                ? "Atribuído"
                : "Pendente"
          }</td>
          <td><span class="status-badge status-${escapeHtml(s.status)}">${escapeHtml(formatarStatus(s.status))}</span></td>
          <td>${dataFormatada}</td>
          <td>${botoesAcoes}</td>
        `;

        tbodySolicitacoes.appendChild(row);
      }
    }

    atualizarEstatisticas();
  }

  // ================= DETALHES / AÇÕES =================
  function verDetalhes(id) {
    const solicitacao = solicitacoes.find((s) => s.id === id);
    if (!solicitacao || !conteudoDetalhes) return;

    const dataSolicitacao = formatDateBR(solicitacao.dataSolicitacao);
    const dataCriacao = solicitacao.dataCriacao ? new Date(solicitacao.dataCriacao).toLocaleString("pt-BR") : "—";
    const dataConclusaoPrevista = formatDateBR(solicitacao.dataConclusaoPrevista);
    const dataConclusaoReal = solicitacao.dataConclusaoReal ? formatDateBR(solicitacao.dataConclusaoReal) : "Não concluída";
    const dataEntrega = solicitacao.dataEntrega ? formatDateBR(solicitacao.dataEntrega) : "Não informado";

    let statusPrazo = "";
    if (solicitacao.status === "finalizado" && solicitacao.dataConclusaoReal) {
      const prevista = parseDateOnly(solicitacao.dataConclusaoPrevista);
      const real = parseDateOnly(solicitacao.dataConclusaoReal);
      if (prevista && real) {
        statusPrazo =
          real <= prevista
            ? `<span style="color: var(--success); font-weight: 800;">✅ Dentro do prazo</span>`
            : `<span style="color: var(--danger); font-weight: 800;">❌ Fora do prazo</span>`;
      }
    }

    const produtos = solicitacao.produtos || {};
    let produtosHtml = "";
    if (produtos.mapa) produtosHtml += "✅ Mapa / Planta<br>";
    if (produtos.croqui) produtosHtml += "✅ Croqui<br>";
    if (produtos.shapefile) produtosHtml += "✅ Shapefile (SHP)<br>";
    if (produtos.kml) produtosHtml += "✅ KMZ/KML<br>";

    const elementos = solicitacao.elementos || {};
    let elementosHtml = "";
    if (elementos.localizacao) elementosHtml += "✅ Localização<br>";
    if (elementos.acessoLocal) elementosHtml += "✅ Via de Acesso Local<br>";
    if (elementos.acessoRegional) elementosHtml += "✅ Via de Acesso Regional<br>";
    if (elementos.areaAmostral) elementosHtml += "✅ Área Amostral<br>";
    if (elementos.outros && elementos.outrosTexto) {
      elementosHtml += `✅ Outros: ${escapeHtml(elementos.outrosTexto)}<br>`;
    }

    conteudoDetalhes.innerHTML = `
      <div class="detalhe-grid">
        <div class="detalhe-item">
          <div class="detalhe-label">ID</div>
          <div class="detalhe-value">#${String(solicitacao.id).padStart(4, "0")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Solicitante</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.solicitante || "")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Cliente</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.cliente || "")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Estudo</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.nomeEstudo || "Não informado")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Empreendimento</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.empreendimento || "Não informado")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Localidade</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.localidade || "")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Município/Estado</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.municipio || "Não informado")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Data de Entrega</div>
          <div class="detalhe-value">${dataEntrega}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Tipo de Mapa</div>
          <div class="detalhe-value">${escapeHtml(formatarTipoMapa(solicitacao.tipoMapa, solicitacao.nomeTipoMapa))}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Diretório dos Arquivos</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.diretorioArquivos || "Não especificado")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Diretório de Salvamento</div>
          <div class="detalhe-value">${escapeHtml(solicitacao.diretorioSalvamento || "Não especificado")}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">ART</div>
          <div class="detalhe-value">${solicitacao.artNecessaria === "sim" ? "✅ Sim" : "❌ Não"}</div>
        </div>

        ${
          solicitacao.artNecessaria === "sim" && solicitacao.artResponsavel
            ? `
          <div class="detalhe-item">
            <div class="detalhe-label">Responsável Técnico</div>
            <div class="detalhe-value">${escapeHtml(solicitacao.artResponsavel)}</div>
          </div>
        `
            : ""
        }

        <div class="detalhe-item">
          <div class="detalhe-label">Finalidade</div>
          <div class="detalhe-value">${escapeHtml(formatarFinalidade(solicitacao.finalidade))}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Status</div>
          <div class="detalhe-value">
            <span class="status-badge status-${escapeHtml(solicitacao.status)}">${escapeHtml(formatarStatus(solicitacao.status))}</span>
          </div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Técnico</div>
          <div class="detalhe-value">${escapeHtml(formatarNomeTecnico(solicitacao.tecnicoResponsavel))}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Prazo Solicitado</div>
          <div class="detalhe-value">${Number(solicitacao.prazoDias || 0)} dias úteis</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Conclusão Prevista</div>
          <div class="detalhe-value">${dataConclusaoPrevista}</div>
        </div>

        <div class="detalhe-item">
          <div class="detalhe-label">Conclusão Real</div>
          <div class="detalhe-value">${dataConclusaoReal}</div>
        </div>

        ${
          statusPrazo
            ? `
          <div class="detalhe-item">
            <div class="detalhe-label">Prazo</div>
            <div class="detalhe-value">${statusPrazo}</div>
          </div>
        `
            : ""
        }

        <div class="detalhe-item">
          <div class="detalhe-label">Data da Solicitação</div>
          <div class="detalhe-value">${dataSolicitacao}</div>
        </div>
      </div>

      <div style="margin-top: 18px;">
        <h4 style="margin-bottom: 10px; font-weight: 900;">📦 Produtos Solicitados</h4>
        <div style="background: rgba(255,255,255,0.04); border: 1px solid var(--stroke); padding: 14px; border-radius: 14px;">
          ${produtosHtml || "Nenhum produto selecionado"}
        </div>
      </div>

      <div style="margin-top: 14px;">
        <h4 style="margin-bottom: 10px; font-weight: 900;">✏️ Croquis</h4>
        <div style="background: rgba(255,255,255,0.04); border: 1px solid var(--stroke); padding: 14px; border-radius: 14px;">
          ${elementosHtml || "Nenhum elemento selecionado"}
        </div>
      </div>

      ${
        solicitacao.observacoes
          ? `
        <div style="margin-top: 14px;">
          <h4 style="margin-bottom: 10px; font-weight: 900;">📝 Observações</h4>
          <div style="background: rgba(255,255,255,0.04); border: 1px solid var(--stroke); padding: 14px; border-radius: 14px; white-space: pre-wrap;">
            ${escapeHtml(solicitacao.observacoes)}
          </div>
        </div>
      `
          : ""
      }

      <div style="margin-top: 16px; padding-top: 12px; border-top: 1px solid var(--stroke);">
        <p style="color: var(--muted); font-size: 0.9rem;">Criado em: ${escapeHtml(dataCriacao)}</p>
      </div>

      <div style="margin-top: 16px; display: flex; gap: 10px; flex-wrap: wrap; justify-content: center;">
        ${
          acessoGestor
            ? `
          <button class="btn btn-warning" type="button" onclick="mudarStatus(${solicitacao.id}, 'processando')">⚙️ Processando</button>
          <button class="btn btn-info" type="button" onclick="mudarStatus(${solicitacao.id}, 'aguardando')">⏳ Aguardando</button>
          <button class="btn btn-success" type="button" onclick="finalizarSolicitacao(${solicitacao.id})">✅ Finalizar</button>
          <button class="btn btn-primary" type="button" onclick="abrirModalAtribuicao(${solicitacao.id})">👤 Atribuir</button>
          <button class="btn btn-danger" type="button" onclick="confirmarExclusao(${solicitacao.id})">🗑️ Excluir</button>
        `
            : acessoTecnico && solicitacao.tecnicoResponsavel === tecnicoLogado
              ? `
          <button class="btn btn-warning" type="button" onclick="mudarStatus(${solicitacao.id}, 'processando')">⚙️ Processando</button>
          <button class="btn btn-info" type="button" onclick="mudarStatus(${solicitacao.id}, 'aguardando')">⏳ Aguardando</button>
          <button class="btn btn-success" type="button" onclick="finalizarSolicitacao(${solicitacao.id})">✅ Finalizar</button>
        `
              : `
          <button class="btn btn-ghost" type="button" disabled style="opacity:.8; cursor:not-allowed;">
            ${solicitacao.status === "finalizado" ? "✅ Finalizada" : "⏳ Em andamento"}
          </button>
        `
        }
        <button class="btn btn-ghost" type="button" onclick="fecharModal()">Fechar</button>
      </div>
    `;

    modalDetalhes?.classList.add("active");
  }

  function fecharModal() {
    modalDetalhes?.classList.remove("active");
  }

  function mudarStatus(id, novoStatus) {
    if (!acessoGestor && !acessoTecnico) {
      mostrarNotificacao("Faça login para alterar status!", "warning");
      return;
    }

    const solicitacao = solicitacoes.find((s) => s.id === id);
    if (!solicitacao) return;

    if (acessoTecnico && solicitacao.tecnicoResponsavel !== tecnicoLogado) {
      mostrarNotificacao("Você só pode alterar solicitações atribuídas a você!", "warning");
      return;
    }

    atualizarSolicitacaoFirebase(id, { status: novoStatus });

    const mensagem = {
      fila: "movida para a fila",
      processando: "iniciou o processamento",
      aguardando: "está aguardando dados",
      finalizado: "foi finalizada",
    };

    mostrarNotificacao(`Solicitação #${id} ${mensagem[novoStatus] || "atualizada"}!`, "success");
  }

  function finalizarSolicitacao(id) {
    if (!acessoGestor && !acessoTecnico) {
      mostrarNotificacao("Faça login para finalizar!", "warning");
      return;
    }

    const solicitacao = solicitacoes.find((s) => s.id === id);
    if (!solicitacao) return;

    if (acessoTecnico && solicitacao.tecnicoResponsavel !== tecnicoLogado) {
      mostrarNotificacao("Você só pode finalizar solicitações atribuídas a você!", "warning");
      return;
    }

    if (solicitacao.status === "finalizado") {
      mostrarNotificacao("Esta solicitação já foi finalizada.", "info");
      return;
    }

    atualizarSolicitacaoFirebase(id, {
      status: "finalizado",
      dataConclusaoReal: new Date().toISOString().split("T")[0],
    });

    fecharModal();
    mostrarNotificacao(`✅ Solicitação #${id} finalizada!`, "success");
  }

  function abrirModalAtribuicao(id) {
    if (!acessoGestor) {
      abrirModalAcessoGestor();
      return;
    }

    const solicitacao = solicitacoes.find((s) => s.id === id);
    if (!solicitacao || !conteudoAtribuicao) return;

    conteudoAtribuicao.innerHTML = `
      <p style="font-size: 1.05rem; margin-bottom: 14px;">
        <strong>Atribuir Técnico - Solicitação #${id}</strong>
      </p>
      <p style="margin-bottom: 14px; color: var(--muted);">
        <strong>Estudo:</strong> ${escapeHtml(solicitacao.nomeEstudo || "Não informado")}<br>
        <strong>Cliente:</strong> ${escapeHtml(solicitacao.cliente || "")}
      </p>

      <div class="form-group">
        <label for="selectTecnico">Selecione o Técnico *</label>
        <select id="selectTecnico">
          <option value="PENDENTE">----------</option>
          <option value="LAIS">Laís</option>
          <option value="LAIZE">Laize</option>
          <option value="VALESKA">Valeska</option>
          <option value="LIZABETH">Lizabeth</option>
          <option value="ISMAEL">Ismael</option>
          <option value="FERNANDO">Fernando</option>
        </select>
      </div>

      <div class="form-group">
        <label for="inputDataConclusao">Data de Conclusão (Prazo)</label>
        <input type="date" id="inputDataConclusao" value="${escapeAttr(solicitacao.dataConclusaoPrevista || "")}">
      </div>

      <div style="display: flex; gap: 10px; justify-content: flex-end; margin-top: 14px;">
        <button class="btn btn-ghost" type="button" onclick="fecharModalAtribuicao()">Cancelar</button>
        <button class="btn btn-primary" type="button" onclick="atribuirTecnico(${id})">Atribuir</button>
      </div>
    `;

    const sel = document.getElementById("selectTecnico");
    if (sel && solicitacao.tecnicoResponsavel) sel.value = solicitacao.tecnicoResponsavel;

    modalAtribuicao?.classList.add("active");
  }

  function fecharModalAtribuicao() {
    modalAtribuicao?.classList.remove("active");
  }

  function atribuirTecnico(id) {
    const tecnico = document.getElementById("selectTecnico")?.value;
    const dataConclusao = document.getElementById("inputDataConclusao")?.value;

    if (!tecnico || tecnico === "PENDENTE") {
      mostrarNotificacao("Selecione um técnico!", "warning");
      return;
    }

    atualizarSolicitacaoFirebase(id, {
      tecnicoResponsavel: tecnico,
      dataConclusaoPrevista: dataConclusao || null,
      status: "processando",
    });

    fecharModalAtribuicao();
    mostrarNotificacao(`Solicitação #${id} atribuída para ${formatarNomeTecnico(tecnico)}!`, "success");
  }

  function confirmarExclusao(id) {
    if (!acessoGestor) {
      abrirModalAcessoGestor();
      return;
    }

    const solicitacao = solicitacoes.find((s) => s.id === id);
    if (!solicitacao || !conteudoConfirmacao) return;

    conteudoConfirmacao.innerHTML = `
      <p style="font-size: 1.05rem; margin-bottom: 12px;">
        🗑️ <strong>Confirmar Exclusão</strong>
      </p>
      <p style="margin-bottom: 12px; color: var(--muted);">
        Tem certeza que deseja excluir a solicitação <strong>#${id}</strong> de <strong>${escapeHtml(solicitacao.cliente || "")}</strong>?
      </p>
      <p style="color: var(--danger); font-weight: 900; margin-bottom: 14px;">
        ⚠️ Esta ação não pode ser desfeita!
      </p>
      <div style="display: flex; gap: 10px; justify-content: flex-end;">
        <button class="btn btn-ghost" type="button" onclick="fecharModalConfirmacao()">Cancelar</button>
        <button class="btn btn-danger" type="button" onclick="excluirSolicitacao(${id})">Excluir</button>
      </div>
    `;

    modalConfirmacao?.classList.add("active");
  }

  function fecharModalConfirmacao() {
    modalConfirmacao?.classList.remove("active");
  }

  // ================= RELATÓRIOS / FILTROS =================
  function filtrarSolicitacoes(tipo) {
    let statusFiltro = null;
    switch (tipo) {
      case "em_andamento":
        statusFiltro = null;
        break;
      case "na_fila":
        statusFiltro = "fila";
        break;
      case "aguardando_dados":
        statusFiltro = "aguardando";
        break;
      default:
        statusFiltro = null;
    }

    tabBtns?.forEach((b) => b.classList.remove("active"));
    atualizarTabela(statusFiltro);
  }

  function filtrarMinhasSolicitacoes(tipo) {
    let statusFiltro = null;
    switch (tipo) {
      case "processando":
        statusFiltro = "processando";
        break;
      case "aguardando":
        statusFiltro = "aguardando";
        break;
      case "finalizadas":
      case "finalizado":
        statusFiltro = "finalizado";
        break;
      default:
        statusFiltro = null;
    }

    tabBtns?.forEach((b) => b.classList.remove("active"));
    atualizarTabela(statusFiltro);
  }

  function abrirModalRelatorio() {
    if (!acessoGestor) {
      abrirModalAcessoGestor();
      return;
    }
    modalRelatorio?.classList.add("active");
  }

  function fecharModalRelatorio() {
    modalRelatorio?.classList.remove("active");
    const resultado = document.getElementById("resultadoRelatorio");
    if (resultado) {
      resultado.style.display = "none";
      resultado.innerHTML = "";
    }
  }

  function gerarRelatorio() {
    if (!acessoGestor) {
      abrirModalAcessoGestor();
      return;
    }

    const dataInicio = document.getElementById("relatorioPeriodoInicio")?.value;
    const dataFim = document.getElementById("relatorioPeriodoFim")?.value;

    if (!dataInicio || !dataFim) {
      mostrarNotificacao("Selecione o período!", "warning");
      return;
    }
    if (parseDateOnly(dataInicio) > parseDateOnly(dataFim)) {
      mostrarNotificacao("Data inicial maior que final!", "error");
      return;
    }

    const solicitacoesPeriodo = solicitacoes.filter((s) => {
      const d = parseDateOnly(s.dataSolicitacao);
      return d && d >= parseDateOnly(dataInicio) && d <= parseDateOnly(dataFim);
    });

    const total = solicitacoesPeriodo.length;
    const finalizadas = solicitacoesPeriodo.filter((s) => s.status === "finalizado").length;

    const dentroPrazo = solicitacoesPeriodo.filter((s) => {
      const real = s.dataConclusaoReal ? parseDateOnly(s.dataConclusaoReal) : null;
      const prevista = s.dataConclusaoPrevista ? parseDateOnly(s.dataConclusaoPrevista) : null;
      return s.status === "finalizado" && real && prevista && real <= prevista;
    }).length;

    const foraPrazo = finalizadas - dentroPrazo;

    const tecnicos = Object.keys(CODIGOS_TECNICOS);
    const estatisticasTecnicos = {};

    tecnicos.forEach((tecnico) => {
      const solTec = solicitacoesPeriodo.filter((s) => s.tecnicoResponsavel === tecnico);
      estatisticasTecnicos[tecnico] = {
        total: solTec.length,
        finalizadas: solTec.filter((s) => s.status === "finalizado").length,
        dentroPrazo: solTec.filter((s) => {
          const real = s.dataConclusaoReal ? parseDateOnly(s.dataConclusaoReal) : null;
          const prevista = s.dataConclusaoPrevista ? parseDateOnly(s.dataConclusaoPrevista) : null;
          return s.status === "finalizado" && real && prevista && real <= prevista;
        }).length,
      };
    });

    let htmlRelatorio = `
      <div class="relatorio-item">
        <h4>📊 Período</h4>
        <p><strong>De:</strong> ${formatDateBR(dataInicio)}</p>
        <p><strong>Até:</strong> ${formatDateBR(dataFim)}</p>
      </div>

      <div class="relatorio-item">
        <h4>📈 Métricas Gerais</h4>
        <div class="relatorio-grid">
          <div class="relatorio-metric">
            <div class="relatorio-metric-value">${total}</div>
            <div class="relatorio-metric-label">Total</div>
          </div>
          <div class="relatorio-metric">
            <div class="relatorio-metric-value">${finalizadas}</div>
            <div class="relatorio-metric-label">Finalizadas</div>
          </div>
          <div class="relatorio-metric">
            <div class="relatorio-metric-value" style="color: var(--success);">${dentroPrazo}</div>
            <div class="relatorio-metric-label">Dentro do Prazo</div>
          </div>
          <div class="relatorio-metric">
            <div class="relatorio-metric-value" style="color: var(--danger);">${foraPrazo}</div>
            <div class="relatorio-metric-label">Fora do Prazo</div>
          </div>
        </div>
      </div>

      <div class="relatorio-item">
        <h4>👨‍💼 Por Técnico</h4>
    `;

    tecnicos.forEach((tecnico) => {
      const stats = estatisticasTecnicos[tecnico];
      if (stats.total > 0) {
        const taxaConclusao = ((stats.finalizadas / stats.total) * 100).toFixed(1);
        const taxaPrazo = stats.finalizadas > 0 ? ((stats.dentroPrazo / stats.finalizadas) * 100).toFixed(1) : "0.0";

        htmlRelatorio += `
          <div class="tecnico-relatorio">
            <h5>${formatarNomeTecnico(tecnico)}</h5>
            <div class="relatorio-grid">
              <div class="relatorio-metric">
                <div class="relatorio-metric-value">${stats.total}</div>
                <div class="relatorio-metric-label">Atribuídas</div>
              </div>
              <div class="relatorio-metric">
                <div class="relatorio-metric-value">${stats.finalizadas}</div>
                <div class="relatorio-metric-label">Finalizadas</div>
              </div>
              <div class="relatorio-metric">
                <div class="relatorio-metric-value">${taxaConclusao}%</div>
                <div class="relatorio-metric-label">Conclusão</div>
              </div>
              <div class="relatorio-metric">
                <div class="relatorio-metric-value">${taxaPrazo}%</div>
                <div class="relatorio-metric-label">No Prazo</div>
              </div>
            </div>
          </div>
        `;
      }
    });

    htmlRelatorio += `
      </div>

      <div style="margin-top: 14px; display:flex; gap:10px; justify-content:flex-end;">
        <button class="btn btn-ghost" type="button" onclick="fecharModalRelatorio()">Fechar</button>
        <button class="btn btn-warning" type="button" onclick="exportarRelatorioCompleto('${escapeAttr(dataInicio)}', '${escapeAttr(dataFim)}')">📥 Exportar CSV</button>
      </div>
    `;

    const resultado = document.getElementById("resultadoRelatorio");
    if (resultado) {
      resultado.style.display = "block";
      resultado.innerHTML = htmlRelatorio;
    }
  }

  function exportarRelatorioCompleto(dataInicio, dataFim) {
    if (!acessoGestor) return;

    const inicio = parseDateOnly(dataInicio);
    const fim = parseDateOnly(dataFim);

    const solicitacoesPeriodo = solicitacoes.filter((s) => {
      const d = parseDateOnly(s.dataSolicitacao);
      return d && inicio && fim && d >= inicio && d <= fim;
    });

    const dadosExport = solicitacoesPeriodo.map((s) => {
      const dentroPrazo =
        s.status === "finalizado" &&
        s.dataConclusaoReal &&
        s.dataConclusaoPrevista &&
        parseDateOnly(s.dataConclusaoReal) <= parseDateOnly(s.dataConclusaoPrevista);

      return {
        ID: s.id,
        Solicitante: s.solicitante,
        Cliente: s.cliente,
        Estudo: s.nomeEstudo || "Não informado",
        "Tipo Mapa": formatarTipoMapa(s.tipoMapa, s.nomeTipoMapa),
        Município: s.municipio,
        Técnico: formatarNomeTecnico(s.tecnicoResponsavel),
        Status: formatarStatus(s.status),
        "Prazo (dias úteis)": s.prazoDias,
        "Diretório Arquivos": s.diretorioArquivos || "Não especificado",
        "Diretório Salvamento": s.diretorioSalvamento || "Não especificado",
        "Data Solicitação": formatDateBR(s.dataSolicitacao),
        "Conclusão Prevista": formatDateBR(s.dataConclusaoPrevista),
        "Conclusão Real": s.dataConclusaoReal ? formatDateBR(s.dataConclusaoReal) : "",
        "Dentro do Prazo": s.status === "finalizado" ? (dentroPrazo ? "SIM" : "NÃO") : "EM ANDAMENTO",
      };
    });

    downloadCSV(dadosExport, `relatorio_solicitacoes_${dataInicio}_a_${dataFim}.csv`);
    mostrarNotificacao("Relatório exportado!", "success");
  }

  function exportarTodosDados() {
    if (!acessoGestor) {
      abrirModalAcessoGestor();
      return;
    }

    const dadosExport = solicitacoes.map((s) => {
      const dentroPrazo =
        s.status === "finalizado" &&
        s.dataConclusaoReal &&
        s.dataConclusaoPrevista &&
        parseDateOnly(s.dataConclusaoReal) <= parseDateOnly(s.dataConclusaoPrevista);

      return {
        ID: s.id,
        Solicitante: s.solicitante,
        Cliente: s.cliente,
        Estudo: s.nomeEstudo || "Não informado",
        "Tipo Mapa": formatarTipoMapa(s.tipoMapa, s.nomeTipoMapa),
        Município: s.municipio,
        Técnico: formatarNomeTecnico(s.tecnicoResponsavel),
        Status: formatarStatus(s.status),
        "Prazo (dias úteis)": s.prazoDias,
        "Diretório Arquivos": s.diretorioArquivos || "Não especificado",
        "Diretório Salvamento": s.diretorioSalvamento || "Não especificado",
        "Data Solicitação": formatDateBR(s.dataSolicitacao),
        "Conclusão Prevista": formatDateBR(s.dataConclusaoPrevista),
        "Conclusão Real": s.dataConclusaoReal ? formatDateBR(s.dataConclusaoReal) : "",
        "Dentro do Prazo": s.status === "finalizado" ? (dentroPrazo ? "SIM" : "NÃO") : "EM ANDAMENTO",
      };
    });

    downloadCSV(dadosExport, `todos_dados_solicitacoes_${new Date().toISOString().split("T")[0]}.csv`);
    mostrarNotificacao("Todos os dados exportados!", "success");
  }

  // ================= ESTATÍSTICAS =================
  function atualizarEstatisticas() {
    const base =
      acessoTecnico && tecnicoLogado
        ? solicitacoes.filter((s) => s.tecnicoResponsavel === tecnicoLogado)
        : solicitacoes;

    setText("totalSolicitacoes", base.length);
    setText("totalFila", base.filter((s) => s.status === "fila").length);
    setText("totalProcessando", base.filter((s) => s.status === "processando").length);
    setText("totalFinalizadas", base.filter((s) => s.status === "finalizado").length);
  }

  function atualizarListaTecnicos() {
    if (!listaTecnicos) return;

    if (!acessoGestor) {
      listaTecnicos.innerHTML =
        '<p style="color: var(--muted); text-align:center; font-size:.9rem;">Faça login como gestor para ver os técnicos</p>';
      return;
    }

    const tecnicos = Object.keys(CODIGOS_TECNICOS);
    let html = "";

    tecnicos.forEach((tecnico) => {
      const solTec = solicitacoes.filter((s) => s.tecnicoResponsavel === tecnico && s.status !== "finalizado");
      if (solTec.length > 0) {
        html += `
          <div class="tecnico-item">
            <strong>${formatarNomeTecnico(tecnico)}</strong>
            <span style="color: var(--muted);">${solTec.length} projeto(s) em andamento</span>
            <ul style="margin-top: 8px; margin-left: 18px; font-size: 0.85rem; padding-left: 14px;">
        `;

        solTec.slice(0, 3).forEach((s) => {
          html += `<li>${escapeHtml(s.nomeEstudo || "Sem nome")} - <span style="color: var(--warning);">${escapeHtml(
            formatarStatus(s.status)
          )}</span></li>`;
        });

        if (solTec.length > 3) {
          html += `<li style="color: var(--muted);">+ ${solTec.length - 3} outro(s)</li>`;
        }

        html += `</ul></div>`;
      }
    });

    if (!html) {
      html =
        '<p style="color: var(--muted); text-align:center; padding:10px; font-size:.9rem;">Nenhum técnico com projetos em andamento</p>';
    }

    listaTecnicos.innerHTML = html;
  }

  function atualizarEstatisticasTecnico() {
    if (!estatisticasTecnico) return;

    if (!acessoTecnico || !tecnicoLogado) {
      estatisticasTecnico.innerHTML =
        '<p style="color: var(--muted); text-align:center; font-size:.9rem;">Faça login como técnico para ver suas estatísticas</p>';
      return;
    }

    const solTec = solicitacoes.filter((s) => s.tecnicoResponsavel === tecnicoLogado);
    const emAndamento = solTec.filter((s) => s.status !== "finalizado").length;
    const finalizadas = solTec.filter((s) => s.status === "finalizado").length;

    const dentroPrazo = solTec.filter((s) => {
      const real = s.dataConclusaoReal ? parseDateOnly(s.dataConclusaoReal) : null;
      const prevista = s.dataConclusaoPrevista ? parseDateOnly(s.dataConclusaoPrevista) : null;
      return s.status === "finalizado" && real && prevista && real <= prevista;
    }).length;

    estatisticasTecnico.innerHTML = `
      <div class="estatistica-item">
        <strong>Total</strong>
        <span>${solTec.length}</span>
      </div>
      <div class="estatistica-item">
        <strong>Em Andamento</strong>
        <span>${emAndamento}</span>
      </div>
      <div class="estatistica-item">
        <strong>Finalizadas</strong>
        <span>${finalizadas}</span>
      </div>
      <div class="estatistica-item">
        <strong>Dentro do Prazo</strong>
        <span style="color: var(--success); font-weight: 800;">${dentroPrazo}</span>
      </div>
    `;
  }

  // ================= HELPERS =================
  function bindEnterNoModal(modalEl, callback) {
    if (!modalEl) return;
    modalEl.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      callback();
    });
  }

  function parseDateOnly(yyyy_mm_dd) {
    if (!yyyy_mm_dd) return null;
    const d = new Date(`${yyyy_mm_dd}T00:00:00`);
    return isNaN(d) ? null : d;
  }

  function formatDateBR(dateOrStr) {
    const d = typeof dateOrStr === "string" ? parseDateOnly(dateOrStr) : dateOrStr;
    return d instanceof Date && !isNaN(d) ? d.toLocaleDateString("pt-BR") : "—";
  }

  function formatarNomeTecnico(codigo) {
    return (
      {
        LAIS: "Laís",
        LAIZE: "Laize",
        VALESKA: "Valeska",
        LIZABETH: "Lizabeth",
        ISMAEL: "Ismael",
        FERNANDO: "Fernando",
        PENDENTE: "Pendente",
      }[codigo] || codigo
    );
  }

  function formatarStatus(status) {
    return (
      {
        fila: "Na Fila",
        processando: "Processando",
        aguardando: "Aguardando Dados",
        finalizado: "Finalizado",
      }[status] || status
    );
  }

  function formatarTipoMapa(tipo, nomeTipoMapa = "") {
    if (tipo === "OUTROS") return nomeTipoMapa?.trim() ? nomeTipoMapa.trim() : "Outros";

    const tipos = {
      TOPOGRAFICO: "Topográfico",
      GEOLOGICO: "Geológico",
      PEDOLOGICO: "Pedológico",
      HIDROGRAFICO: "Hidrográfico",
      VEGETACAO: "Vegetação",
      FAUNA: "Fauna",
      FLORA: "Flora",
      PLANTA_GEORREFERENCIADA: "Planta Georreferenciada",
      USO_SOLO: "Uso do Solo",
      GEOMORFOLOGICO: "Geomorfológico",
      LOCALIZACAO: "Localização",
    };
    return tipos[tipo] || tipo;
  }

  function formatarFinalidade(finalidade) {
    return (
      {
        LICENCIAMENTO: "Licenciamento Ambiental",
        ESTUDO_IMPACTO: "Estudo de Impacto Ambiental",
        PRE_CAMPO: "Pré-Campo",
        POS_CAMPO: "Pós-Campo",
        PLANEJAMENTO: "Planejamento Territorial",
        OUTROS: "Outros",
      }[finalidade] || finalidade
    );
  }

  function calcularDataConclusao(dataInicio, prazoDias) {
    const data = parseDateOnly(dataInicio);
    if (!data) return null;

    let dias = 0;
    while (dias < prazoDias) {
      data.setDate(data.getDate() + 1);
      if (data.getDay() !== 0 && data.getDay() !== 6) dias++;
    }
    return data.toISOString().split("T")[0];
  }

  function setInputDate(id, date) {
    const el = document.getElementById(id);
    if (el && date instanceof Date && !isNaN(date)) el.value = date.toISOString().split("T")[0];
  }

  function safeSetDisplay(id, val) {
    const el = document.getElementById(id);
    if (el) el.style.display = val;
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = String(value);
  }

  function mostrarNotificacao(mensagem, tipo = "info") {
    const n = document.createElement("div");
    n.className = `notification ${tipo}`;
    n.innerHTML = `<span>${getEmoji(tipo)}</span><span>${escapeHtml(mensagem)}</span>`;
    document.body.appendChild(n);
    setTimeout(() => n.remove(), 3500);
  }

  function getEmoji(tipo) {
    return { success: "✅", error: "❌", warning: "⚠️", info: "ℹ️" }[tipo] || "ℹ️";
  }

  function escapeHtml(str) {
    return String(str ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function escapeAttr(str) {
    return escapeHtml(str).replaceAll("\n", " ");
  }

  function convertToCSV(data) {
    if (!data || data.length === 0) return "";
    const headers = Object.keys(data[0]);

    const rows = data.map((obj) =>
      headers
        .map((h) => {
          const v = obj[h] ?? "";
          return `"${String(v).replace(/"/g, '""')}"`;
        })
        .join(",")
    );

    return [headers.join(","), ...rows].join("\n");
  }

  function downloadCSV(dataArray, filename) {
    const csv = convertToCSV(dataArray);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    link.setAttribute("download", filename);
    link.style.visibility = "hidden";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // ================= EXPORT GLOBAL =================
  window.toggleForm = toggleForm;
  window.limparForm = limparForm;
  window.toggleTipoMapaOutros = toggleTipoMapaOutros;
  window.toggleARTResponsavel = toggleARTResponsavel;
  window.toggleElementosOutros = toggleElementosOutros;

  window.abrirModalAcessoGestor = abrirModalAcessoGestor;
  window.fecharModalAcessoGestor = fecharModalAcessoGestor;
  window.validarCodigoAcesso = validarCodigoAcesso;

  window.abrirModalAcessoTecnico = abrirModalAcessoTecnico;
  window.fecharModalAcessoTecnico = fecharModalAcessoTecnico;
  window.validarAcessoTecnico = validarAcessoTecnico;

  window.fazerLogout = fazerLogout;

  window.atualizarTabela = atualizarTabela;
  window.verDetalhes = verDetalhes;
  window.fecharModal = fecharModal;

  window.abrirModalAtribuicao = abrirModalAtribuicao;
  window.fecharModalAtribuicao = fecharModalAtribuicao;
  window.atribuirTecnico = atribuirTecnico;

  window.mudarStatus = mudarStatus;
  window.finalizarSolicitacao = finalizarSolicitacao;

  window.confirmarExclusao = confirmarExclusao;
  window.fecharModalConfirmacao = fecharModalConfirmacao;
  window.excluirSolicitacao = excluirSolicitacao;

  window.abrirModalRelatorio = abrirModalRelatorio;
  window.fecharModalRelatorio = fecharModalRelatorio;
  window.gerarRelatorio = gerarRelatorio;
  window.exportarRelatorioCompleto = exportarRelatorioCompleto;
  window.exportarTodosDados = exportarTodosDados;

  window.filtrarSolicitacoes = filtrarSolicitacoes;
  window.filtrarMinhasSolicitacoes = filtrarMinhasSolicitacoes;
})();