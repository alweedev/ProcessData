// @ts-check
import { test, expect } from "@playwright/test";
import * as XLSX from "xlsx";
import { avancarAteGerar, xlsxFile, validCpf } from "./fixtures.mjs";

// Nomes fictícios sem número e de divisão inequívoca: o que tem dúvida cai na conferência de nomes e trava o Gerar.
const NOMES = ["Ana Souza", "Bruno Lima", "Carla Dias", "Diego Rocha", "Elisa Pinto", "Fabio Moura", "Gisele Nunes"];

/** Ficha completa: traz os 8 campos obrigatórios (CPF, empresa, centro de custo código/descrição,
 *  nome completo, e-mail, telefone e data de nascimento). */
const cadastroRows = (n) =>
  Array.from({ length: n }, (_, i) => ({
    CPF: validCpf(i + 1),
    "NOME COMPLETO": NOMES[i % NOMES.length],
    EMAIL: `p${i + 1}@x.com`,
    EMPRESA: "Empresa A",
    "Centro de custo": "CC1",
    "Descrição Centro de Custo": "ADMINISTRATIVO",
    TELEFONE: `1199999${String(i + 1).padStart(4, "0")}`,
    "Data de Nascimento": "12/05/1990",
    "SOLICITANTE? (S/N)": "S",
  }));

const planilha = (n = 1) => xlsxFile("cadastro.xlsx", cadastroRows(n));

/** Botões da linha do tempo (só os pontos ao alcance viram botões). */
const pontosClicaveis = (page) => page.getByRole("list", { name: "Etapas do cadastro" }).getByRole("button");
const centroX = async (locator) => {
  const box = await locator.boundingBox();
  return box.x + box.width / 2;
};

test.beforeEach(async ({ page }) => {
  await page.goto("/");
  await page.locator("#cadastro-tab").click();
  await expect(page.locator("#cadastro")).toBeVisible(); // painel da aba
  await expect(page.locator("#cadastro_files")).toBeAttached(); // 1ª etapa: enviar fichas
});

test("gera saida_cadastro.xlsx e limpa a selecao no sucesso", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page);

  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  expect(download.suggestedFilename()).toBe("saida_cadastro.xlsx");

  // A mensagem final descreve o estado real: o cadastro JÁ está concluído (não
  // "o download começou") e cita o mesmo arquivo que o navegador recebeu.
  const status = page.locator("#cadastro_status");
  await expect(status).toContainText("Cadastro concluído");
  await expect(status).toContainText(download.suggestedFilename());
  await expect(status).not.toContainText("começou");
});

test("começa nas fichas e só avança quando há planilha", async ({ page }) => {
  await expect(page.locator("#cadastro_next_btn")).toBeDisabled();
  await expect(page.locator("#cadastro_back_btn")).toHaveCount(0); // não há etapa anterior

  await page.setInputFiles("#cadastro_files", planilha());

  await expect(page.locator("#cadastro_next_btn")).toBeEnabled();
});

test("escolher uma opção avança sozinho para a próxima etapa", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(page.locator("#cadastro_login_choice")).toBeVisible();
  await expect(page.locator("#cadastro_fluxo")).toHaveCount(0);
  // Sem valor padrão: nada vem marcado.
  await expect(page.locator("#cadastro_login_choice-CPF")).toHaveAttribute("aria-checked", "false");

  await page.locator("#cadastro_login_choice-CPF").click();
  await expect(page.locator("#cadastro_fluxo")).toBeVisible();
  await expect(page.locator("#cadastro_fluxo-SELF")).toHaveAttribute("aria-checked", "false");

  await page.locator("#cadastro_fluxo-SELF").click();
  await expect(page.locator("#cadastro_btn")).toBeVisible();
});

test("não dá para pular etapas: só os pontos já alcançados são clicáveis", async ({ page }) => {
  await expect(pontosClicaveis(page)).toHaveCount(0);

  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click(); // agora em "Tipo de login"

  await expect(pontosClicaveis(page)).toHaveCount(1); // só "Fichas"
  await expect(pontosClicaveis(page).first()).toContainText("Fichas");
});

test("Voltar mantém o que já foi escolhido", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();
  await page.locator("#cadastro_login_choice-EMAIL").click(); // avança para o fluxo

  await page.locator("#cadastro_back_btn").click();

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "true");
});

test("clicar num ponto concluído volta à etapa; ao reescolher, vai direto ao fim se o resto já está pronto", async ({
  page,
}) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "CPF", "SELF");

  await pontosClicaveis(page).filter({ hasText: "Tipo de login" }).click();
  await expect(page.locator("#cadastro_login_choice-CPF")).toHaveAttribute("aria-checked", "true");

  await page.locator("#cadastro_login_choice-EMAIL").click();

  // O fluxo já estava escolhido: não faz parar de novo nele.
  await expect(page.locator("#cadastro_btn")).toBeVisible();
});

test("o anel da linha do tempo pula para a etapa atual", async ({ page }) => {
  const anel = page.getByTestId("stepper-marker");
  await expect.poll(async () => Math.abs((await centroX(anel)) - (await centroX(page.getByTestId("stepper-dot-0"))))).toBeLessThan(1);
  await expect(anel.locator("span")).toHaveCSS("animation-name", "none"); // abrir a tela não anima

  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(anel.locator("span")).toHaveCSS("animation-name", "pd-hop"); // o pulo
  await expect.poll(async () => Math.abs((await centroX(anel)) - (await centroX(page.getByTestId("stepper-dot-1"))))).toBeLessThan(1);
});

test("as escolhas não ficam salvas: ao fechar e reabrir é preciso escolher de novo", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "EMAIL", "FRONT");

  await page.reload();
  await page.locator("#cadastro-tab").click();
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();

  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
  await page.locator("#cadastro_login_choice-CPF").click();
  await expect(page.locator("#cadastro_fluxo-FRONT")).toHaveAttribute("aria-checked", "false");
});

test("depois de gerar, 'Novo cadastro' recomeça nas fichas com tudo vazio", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "EMAIL", "FRONT");
  await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  await expect(page.locator("#cadastro_status")).toContainText("Cadastro concluído");
  await expect(pontosClicaveis(page)).toHaveCount(0); // concluído: só "Novo cadastro" recomeça

  await page.locator("#cadastro_new_btn").click();

  await expect(page.locator("#cadastro_files")).toBeAttached();
  await expect(page.locator("#cadastro_next_btn")).toBeDisabled();
  await page.setInputFiles("#cadastro_files", planilha());
  await page.locator("#cadastro_next_btn").click();
  await expect(page.locator("#cadastro_login_choice-EMAIL")).toHaveAttribute("aria-checked", "false");
});

test("bloqueia mais de 5 arquivos no cliente", async ({ page }) => {
  const files = cadastroRows(6).map((r, i) => xlsxFile(`c${i}.xlsx`, [r]));
  await page.setInputFiles("#cadastro_files", files);
  // a guarda do cliente recusa a seleção (aqui não havia nada antes) e mostra toast
  await expect
    .poll(async () => page.locator("#cadastro_files").evaluate((el) => el.files.length))
    .toBe(0);
});

test("uma seleção recusada não apaga as fichas que já estavam escolhidas", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await expect(page.locator("#cadastro_uploadFeedback")).toContainText("cadastro.xlsx");

  await page.setInputFiles("#cadastro_files", cadastroRows(6).map((r, i) => xlsxFile(`c${i}.xlsx`, [r])));

  await expect(page.getByText("Máximo de 5 arquivos por envio.")).toBeVisible();
  await expect(page.locator("#cadastro_uploadFeedback")).toContainText("cadastro.xlsx");
  await expect(page.locator("#cadastro_next_btn")).toBeEnabled();
});

const naoPlanilha = { name: "notas.pdf", mimeType: "application/pdf", buffer: Buffer.from("nao sou planilha") };

test("arquivo que não é planilha é recusado com aviso pelo nome e não entra na lista", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", naoPlanilha);

  await expect(page.getByText('"notas.pdf" não é uma planilha — só .xlsx ou .xls são aceitos.')).toBeVisible();
  await expect(page.locator("#cadastro_next_btn")).toBeDisabled();
});

test("numa seleção mista só a planilha entra; o resto é ignorado com aviso", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", [planilha(), naoPlanilha]);

  await expect(page.getByText('"notas.pdf" foi ignorado — só planilhas .xlsx ou .xls são aceitas.')).toBeVisible();
  await expect(page.locator("#cadastro_uploadFeedback")).toContainText("cadastro.xlsx");
  await expect(page.locator("#cadastro_uploadFeedback")).not.toContainText("notas.pdf");
  await expect(page.locator("#cadastro_next_btn")).toBeEnabled();
});

test("ao voltar, as etapas seguintes já preenchidas ficam tracejadas (guardadas), não verdes de concluídas", async ({
  page,
}) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "CPF", "SELF");
  await expect(page.getByTestId("stepper-dot-1")).not.toHaveClass(/border-dashed/); // concluídas: preenchimento verde

  await pontosClicaveis(page).filter({ hasText: "Fichas" }).click();

  await expect(page.getByTestId("stepper-dot-1")).toHaveClass(/border-dashed/);
  await expect(page.getByTestId("stepper-dot-2")).toHaveClass(/border-dashed/);
  await expect(page.getByTestId("stepper-dot-3")).not.toHaveClass(/border-dashed/); // "Gerar" não tem valor a guardar

  // Sem fichas nada do que vem depois vale: os pontos voltam a pendentes e deixam de ser clicáveis.
  await page.locator("#cadastro_clear_btn").click();
  await expect(page.getByTestId("stepper-dot-1")).not.toHaveClass(/border-dashed/);
  await expect(pontosClicaveis(page)).toHaveCount(0);
});

test("ao chegar em Gerar a planilha é validada sozinha e o veredito já aparece no topo", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha(3));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
  await expect(page.locator("#cadastro_validation_summary")).toContainText("3 cadastros prontos");
});

test("com pendências, gerar pede confirmação: 'Revisar antes' não gera, 'Gerar mesmo assim' gera", async ({ page }) => {
  const semCpfValido = xlsxFile("cadastro.xlsx", [...cadastroRows(2), { ...cadastroRows(1)[0], CPF: "123" }]);
  await page.setInputFiles("#cadastro_files", semCpfValido);
  await avancarAteGerar(page);
  await expect(page.locator("#cadastro_validation_summary")).toContainText("1 de 3 cadastros tem pendência");

  await page.locator("#cadastro_btn").click();
  const dialogo = page.getByRole("dialog", { name: "Gerar mesmo com pendências?" });
  await expect(dialogo).toBeVisible();

  await page.locator("#cadastro_confirm_cancel_btn").click();
  await expect(dialogo).toHaveCount(0);
  await expect(page.locator("#cadastro_status")).not.toContainText("Cadastro concluído");

  await page.locator("#cadastro_btn").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.locator("#cadastro_confirm_generate_btn").click(),
  ]);
  expect(download.suggestedFilename()).toBe("saida_cadastro.xlsx");
  await expect(page.locator("#cadastro_status")).toContainText("Cadastro concluído");
});

test("campo obrigatório em branco invalida a linha, aponta a linha do Excel e o campo no relatório", async ({ page }) => {
  const semTelefone = xlsxFile("cadastro.xlsx", [...cadastroRows(2), { ...cadastroRows(3)[2], TELEFONE: "" }]);
  await page.setInputFiles("#cadastro_files", semTelefone);
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("1 de 3 cadastros tem pendência");
  const problemas = page.locator("#cadastro_problems");
  await expect(problemas).toContainText("Linha 4 · CARLA DIAS — Telefone em branco"); // 3ª linha de dados = linha 4 do Excel
});

test("telefone em branco usa o CELULAR - CONTATO: a ficha é validada e o número sai no arquivo de carga", async ({ page }) => {
  const linha = { ...cadastroRows(1)[0], TELEFONE: "", "CELULAR - CONTATO": "11988887777" };
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [linha]));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
  await expect(page.locator("#cadastro_validation_summary")).not.toContainText("pendência");
});

test("planilha sem nenhuma coluna reconhecida: um cartão só explica o motivo e o Gerar fica travado", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("estranha.xlsx", [{ A: 1, B: 2 }]));
  await avancarAteGerar(page);

  const erro = page.locator("#cadastro_validation_error");
  await expect(erro).toContainText("Não foi possível validar a planilha");
  await expect(erro).toContainText("estranha.xlsx: Não parece uma ficha de cadastro");
  await expect(erro).toContainText("Colunas lidas: A, B");
  await expect(erro).toContainText("O cabeçalho precisa estar na 1ª linha");
  await expect(erro.getByText("Detalhes técnicos")).toBeVisible();
  await expect(page.getByText("não depende")).toHaveCount(0); // a geração falharia pelo mesmo motivo
  await expect(page.locator("#cadastro_btn")).toBeDisabled(); // e o motivo já está na tela: sem 2º cartão de erro
  await expect(page.locator("#cadastro_debug")).toHaveCount(0);
  await expect(page.locator("button", { hasText: "Trocar fichas" })).toHaveCount(0); // "Alterar" na lista já faz isso

  // Trocar a ficha é ali mesmo, na linha Fichas: sem voltar ao começo, e login/fluxo continuam.
  await page.setInputFiles("#cadastro_swap_files", planilha(2));
  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
  await expect(page.locator("#cadastro_validation_error")).toHaveCount(0);
  await expect(page.locator("#cadastro_btn")).toBeEnabled();
  await expect(page.locator("#cadastro_files")).toHaveCount(0); // continua na última etapa
});

test("trocar a ficha na última etapa: mostra o novo nome, revalida e mantém login e fluxo", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [{ ...cadastroRows(1)[0], TELEFONE: "" }]));
  await avancarAteGerar(page, "EMAIL", "SELF");
  await expect(page.locator("#cadastro_validation_summary")).toContainText("pendência");

  await page.setInputFiles("#cadastro_swap_files", xlsxFile("corrigida.xlsx", cadastroRows(2)));

  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo: 2 cadastros prontos");
  const linha = (idAlterar) => page.locator("li", { has: page.locator(idAlterar) });
  await expect(linha("#cadastro_edit_fichas")).toContainText("corrigida.xlsx");
  await expect(linha("#cadastro_edit_login")).toContainText("E-mail");
  await expect(page.locator("#cadastro_btn")).toBeEnabled();
});

test("e-mail com dois endereços na mesma célula é inválido (o login por e-mail precisa de um só)", async ({ page }) => {
  const doisEmails = { ...cadastroRows(1)[0], EMAIL: "um@x.com; dois@x.com" };
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [doisEmails]));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("O cadastro tem pendência");
  await expect(page.locator("#cadastro_problems")).toContainText("E-mail inválido");
});

test("linha repetida é só aviso: é removida do arquivo, sem pedir confirmação para gerar", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [...cadastroRows(2), cadastroRows(1)[0]]));
  await avancarAteGerar(page);
  await expect(page.locator("#cadastro_validation_summary")).toContainText("1 linha duplicada removida");

  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);

  expect(download.suggestedFilename()).toBe("saida_cadastro.xlsx");
  await expect(page.getByRole("dialog")).toHaveCount(0);
});

test("mudar o login descarta a validação: ao voltar a Gerar ela roda de novo", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page, "CPF", "SELF");
  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");

  await pontosClicaveis(page).filter({ hasText: "Tipo de login" }).click();
  const revalidou = page.waitForRequest((r) => r.url().includes("/api/analysis/summary"));
  await page.locator("#cadastro_login_choice-EMAIL").click(); // o fluxo já estava escolhido: vai direto a Gerar

  await revalidou; // pediu uma validação nova, em vez de reaproveitar o relatório do login anterior
  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
});

// ------------------------------------------------------------ conferência de nomes
// Nomes fictícios: só a estrutura importa. "Xanto Zeko Quim Brag" não tem palavra conhecida (divisão ambígua) e cabe
// em 20; o sobrenome longo passa de 20 e nunca é cortado em silêncio.
const DUVIDOSO = "Xanto Zeko Quim Brag";
const SOBRENOME_LONGO = "Anna Luiza Fernandes de Albuquerque Cavalcanti";

const nomeNaFicha = (nomeCompleto, i = 0) => xlsxFile("cadastro.xlsx", [{ ...cadastroRows(1)[0], CPF: validCpf(50 + i), "NOME COMPLETO": nomeCompleto }]);

/** Lê o arquivo gerado (a primeira aba) como linhas de objetos. */
const lerSaida = (download) =>
  download.path().then((p) => XLSX.utils.sheet_to_json(XLSX.readFile(p).Sheets.Cadastro, { defval: "" }));

test("nome que separa com segurança segue sozinho: nenhuma conferência e o Gerar já liberado", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", nomeNaFicha("Maria Clara da Silva Santos"));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
  await expect(page.locator("#cadastro_names_review")).toHaveCount(0);
  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);

  const [linha] = await lerSaida(download);
  expect([linha.Nome, linha.SobreNome, linha.NomeCompleto]).toEqual(["MARIA CLARA", "DA SILVA SANTOS", "MARIA CLARA DA SILVA SANTOS"]);
});

test("nome duvidoso aparece na conferência e o Gerar só libera depois de aceitar", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", nomeNaFicha(DUVIDOSO, 1));
  await avancarAteGerar(page);

  const painel = page.locator("#cadastro_names_review");
  await expect(painel).toContainText("Nomes para conferir (1)");
  await expect(painel).toContainText("XANTO ZEKO QUIM BRAG"); // o nome como no documento
  await expect(painel).toContainText("a divisão entre nome e sobrenome é ambígua");
  await expect(page.locator("#cadastro_btn")).toBeDisabled();

  await page.locator("#cadastro_names_accept_all").click();

  await expect(painel).toContainText("Todos os nomes foram conferidos");
  await expect(page.locator("#cadastro_btn")).toBeEnabled();
  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  const [linha] = await lerSaida(download);
  expect([linha.Nome, linha.SobreNome]).toEqual(["XANTO", "ZEKO QUIM BRAG"]); // a sugestão aceita como veio
});

test("corrigir a divisão à mão vale no arquivo gerado", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", nomeNaFicha("Yara Zelo Quiv Braz", 2));
  await avancarAteGerar(page);

  await page.locator("#cadastro_name_1-2_nome").fill("Yara Zelo");
  await page.locator("#cadastro_name_1-2_sobrenome").fill("Quiv Braz");

  await expect(page.locator("#cadastro_names_review")).toContainText("Todos os nomes foram conferidos");
  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
  const [linha] = await lerSaida(download);
  expect([linha.Nome, linha.SobreNome, linha.NomeCompleto]).toEqual(["YARA ZELO", "QUIV BRAZ", "YARA ZELO QUIV BRAZ"]);
});

test("nome acima de 20 caracteres não é cortado: bloqueia o Gerar, sugere abreviação e o nome completo não muda", async ({
  page,
}) => {
  await page.setInputFiles("#cadastro_files", nomeNaFicha(SOBRENOME_LONGO, 3));
  await avancarAteGerar(page);

  const sobrenome = page.locator("#cadastro_name_1-2_sobrenome");
  await expect(sobrenome).toHaveValue("F DE A CAVALCANTI"); // a sugestão já cabe: nomes do meio viram inicial
  await expect(page.locator("#cadastro_btn")).toBeDisabled();

  await page.locator("#cadastro_name_1-2_confirm").click();
  await expect(page.locator("#cadastro_btn")).toBeEnabled();
  const [download] = await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);

  const [linha] = await lerSaida(download);
  expect([linha.Nome, linha.SobreNome]).toEqual(["ANNA LUIZA", "F DE A CAVALCANTI"]);
  expect(linha.NomeCompleto).toBe("ANNA LUIZA FERNANDES DE ALBUQUERQUE CAVALCANTI"); // o documento segue inteiro
});

test("texto acima de 20 caracteres mostra o contador no vermelho e não deixa confirmar", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", nomeNaFicha(SOBRENOME_LONGO, 4));
  await avancarAteGerar(page);

  await page.locator("#cadastro_name_1-2_sobrenome").fill("FERNANDES DE ALBUQUERQUE CAVALCANTI");

  await expect(page.getByText("35/20: passa do limite")).toBeVisible();
  await expect(page.locator("#cadastro_name_1-2_confirm")).toBeDisabled();
  await expect(page.locator("#cadastro_btn")).toBeDisabled();
});

test("mudar de ficha descarta as decisões: a conferência recomeça para o novo arquivo", async ({ page }) => {
  // Nome só deste teste: o vocabulário aprende com o que outros testes geraram, e um nome já aceito deixa de ser duvidoso.
  await page.setInputFiles("#cadastro_files", nomeNaFicha("Wenzo Trik Pruv Blan", 5));
  await avancarAteGerar(page);
  await page.locator("#cadastro_names_accept_all").click();
  await expect(page.locator("#cadastro_btn")).toBeEnabled();

  await pontosClicaveis(page).filter({ hasText: "Fichas" }).click();
  await page.setInputFiles("#cadastro_files", nomeNaFicha("Wanda Zuk Prin Blot", 6));
  for (let passo = 0; passo < 3; passo++) await page.locator("#cadastro_next_btn").click(); // fichas, login e fluxo já escolhidos

  await expect(page.locator("#cadastro_names_review")).toContainText("1 nome ainda precisa da sua confirmação");
  await expect(page.locator("#cadastro_btn")).toBeDisabled();
});

test("corrigir o mesmo nome duas vezes ensina o vocabulário: da próxima ele não precisa mais de conferência", async ({
  page,
}) => {
  const nome = "Quenzo Vrill Tamb Grose";
  for (let vez = 0; vez < 2; vez++) {
    await page.goto("/");
    await page.locator("#cadastro-tab").click();
    await page.setInputFiles("#cadastro_files", nomeNaFicha(nome, 10 + vez));
    await avancarAteGerar(page);
    await page.locator("#cadastro_name_1-2_nome").fill("Quenzo Vrill"); // correção manual: pesa mais que aceitar
    await page.locator("#cadastro_name_1-2_sobrenome").fill("Tamb Grose");
    await Promise.all([page.waitForEvent("download"), page.locator("#cadastro_btn").click()]);
    await expect(page.locator("#cadastro_status")).toContainText("Cadastro concluído");
  }

  await page.goto("/");
  await page.locator("#cadastro-tab").click();
  await page.setInputFiles("#cadastro_files", nomeNaFicha(nome, 12));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_validation_summary")).toContainText("Tudo certo");
  await expect(page.locator("#cadastro_names_review")).toHaveCount(0); // aprendeu: divisão segura
});

// ------------------------------------------------------- resultado antes de gerar (últ. etapa)

test("tudo certo é uma linha só: sem repetir a mensagem, com os números a um clique", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", planilha(2));
  await avancarAteGerar(page);

  const resumo = page.locator("#cadastro_validation_summary");
  await expect(resumo).toContainText("Tudo certo: 2 cadastros prontos para gerar.");
  await expect(page.locator("#cadastro_problems")).toHaveCount(0);
  await expect(page.getByText("Nenhum problema encontrado")).toHaveCount(0);
  await expect(page.getByText("A validação é informativa")).toHaveCount(0);
  await expect(page.getByText("Inválidas")).toHaveCount(0); // os números só depois de "Ver detalhes"

  await resumo.getByRole("button", { name: "Ver detalhes" }).click();
  await expect(resumo.getByText("Inválidas")).toBeVisible();
});

test("o resultado vem antes dos botões e não há 'Revalidar': a validação roda sozinha e refaz ao trocar a ficha", async ({
  page,
}) => {
  await page.setInputFiles("#cadastro_files", planilha());
  await avancarAteGerar(page);

  const resumo = await page.locator("#cadastro_validation_summary").boundingBox();
  const gerar = await page.locator("#cadastro_btn").boundingBox();
  expect(resumo.y).toBeLessThan(gerar.y);
  await expect(page.locator("#cadastro_validate_btn")).toHaveCount(0); // revalidar a mesma ficha dá o mesmo resultado
});

test("com pendência, quem corrige vê o problema e o passageiro", async ({
  page,
}) => {
  const rows = [...cadastroRows(3), { ...cadastroRows(4)[3], TELEFONE: "" }, { ...cadastroRows(5)[4], TELEFONE: "" }];
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", rows));
  await avancarAteGerar(page);

  const problemas = page.locator("#cadastro_problems");
  await expect(problemas.getByText("Nos dois:")).toBeVisible(); // dito uma vez
  await expect(problemas.getByText("Telefone em branco")).toHaveCount(1);
  await expect(problemas).toContainText("DIEGO ROCHA");
  await expect(problemas).toContainText("ELISA PINTO");

  await expect(page.locator("button", { hasText: "Trocar fichas" })).toHaveCount(0); // "Alterar" na lista já faz isso
});

test("linhas quase vazias não viram uma parede de problemas: uma frase só e o Gerar continua à vista", async ({ page }) => {
  const modelo = cadastroRows(1)[0];
  const vazia = (nome) => ({ ...Object.fromEntries(Object.keys(modelo).map((k) => [k, ""])), "NOME COMPLETO": nome });
  const rows = [...cadastroRows(2), ...["FULANO UM", "FULANO DOIS", "FULANO TRES", "FULANO QUATRO", "FULANO CINCO", "FULANO SEIS"].map(vazia)];
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", rows));
  await avancarAteGerar(page);

  const problemas = page.locator("#cadastro_problems");
  await expect(problemas).toContainText("6 linhas sem nenhum campo obrigatório — parece que não foram preenchidas");
  await expect(problemas.getByText("em branco")).toHaveCount(0); // sem a lista de 7 problemas por passageiro
  const painel = await problemas.boundingBox();
  expect(painel.height).toBeLessThan(330);
  const gerar = await page.locator("#cadastro_btn").boundingBox();
  expect(gerar.y - painel.y).toBeLessThan(450); // o Gerar fica logo abaixo, não a uma tela de distância
});

test("ficha sem a coluna Telefone: a frase da ficha basta, sem repetir 'em branco' por passageiro", async ({ page }) => {
  const semColuna = cadastroRows(3).map(({ TELEFONE, ...resto }) => resto);
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", semColuna));
  await avancarAteGerar(page);

  const problemas = page.locator("#cadastro_problems");
  await expect(problemas).toContainText('A ficha não tem a coluna "Telefone".');
  await expect(problemas.getByText("em branco")).toHaveCount(0);
  await expect(problemas).not.toContainText("ANA SOUZA");
});

test("veredito e problemas ficam num cartão só, sem repetir a contagem", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [...cadastroRows(3), { ...cadastroRows(4)[3], TELEFONE: "" }]));
  await avancarAteGerar(page);

  const resumo = page.locator("#cadastro_validation_summary");
  await expect(resumo.locator("#cadastro_problems")).toBeVisible(); // o detalhe está dentro do cartão do veredito
  await expect(resumo).toContainText("1 de 4 cadastros tem pendência");
  await expect(page.getByText("O que precisa de atenção")).toHaveCount(0); // sem um segundo título com a mesma conta
});

test("com pendência o 'Gerar' segue em destaque (âmbar) e pede confirmação", async ({ page }) => {
  await page.setInputFiles("#cadastro_files", xlsxFile("cadastro.xlsx", [{ ...cadastroRows(1)[0], TELEFONE: "" }]));
  await avancarAteGerar(page);

  await expect(page.locator("#cadastro_btn")).toHaveClass(/bg-warning/); // preenchido, não um botão apagado

  await expect(page.locator("#cadastro_btn")).toBeEnabled();
  await page.locator("#cadastro_btn").click();
  await expect(page.getByRole("dialog", { name: "Gerar mesmo com pendências?" })).toBeVisible();
});
