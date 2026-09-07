(function () {
  function wireAnaliseHelp() {
    const btn = document.getElementById("analise_help_btn");
    if (!btn) return;
    btn.addEventListener("click", () => {
      const html = `
        <div class="text-start">
          <p class="mb-2 fw-semibold">Guia rápido do workspace</p>
          <ol class="ps-2 small">
            <li><strong>1º Passo:</strong> envie a planilha arrastando ou clicando em <em>Selecionar arquivo</em>.</li>
            <li><strong>2º Passo:</strong> acompanhe o processamento incremental e já navegue pela ficha enquanto carrega.</li>
            <li><strong>3º Passo:</strong> use a <strong>busca</strong> para filtrar qualquer coluna instantaneamente.</li>
            <li><strong>4º Passo:</strong> ajuste <strong>quebra de texto</strong>, <strong>densidade</strong> e fixe o resumo conforme sua preferência.</li>
            <li><strong>5º Passo:</strong> exporte o <strong>CSV</strong>, que respeita filtros ativos e limitações para manter a performance.</li>
          </ol>
          <p class="small text-muted mb-0">Tudo acontece no navegador - nenhum dado sensível sai da sua máquina.</p>
        </div>`;
      if (window.Swal?.fire) {
        Swal.fire({
          title: "Workspace Analítico",
          html,
          confirmButtonText: "Continuar",
          width: 540,
        });
      } else {
        alert("Envie o Excel, ajuste preferências e exporte quando quiser.");
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wireAnaliseHelp);
  } else {
    wireAnaliseHelp();
  }
})();
