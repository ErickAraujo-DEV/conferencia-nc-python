import csv
import re
import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

CSV_PATH = "Cred recebido (relatorio para cadastro no SILOMS).csv"

# ─── FUNÇÕES DE LIMPEZA E TRATAMENTO ─────────────────────────────────────────

def clean_int(val):
    val = str(val).strip()
    if not val: return ""
    try:
        return str(int(float(val)))
    except ValueError:
        return val


def clean_fonte(val):
    val = str(val).strip()
    if not val: return ""
    try:
        return str(int(float(val.replace(",", "."))))
    except ValueError:
        return val


def clean_valor_e_tipo(val):
    """Retorna (valor_formatado_positivo, tipo_nota) detectando negativos com sinal - ou parênteses."""
    val_str = str(val).strip()
    if not val_str:
        return "", "C"

    # Verifica se é negativo (tem sinal de menos ou está entre parênteses)
    is_negative = False
    if "-" in val_str or (val_str.startswith("(") and val_str.endswith(")")):
        is_negative = True

    # Remove os parênteses e o sinal de menos para deixar o número "positivo"
    val_str = val_str.replace("(", "").replace(")", "").replace("-", "").strip()

    try:
        if "." in val_str and "," in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        elif "," in val_str:
            val_str = val_str.replace(",", ".")

        f = float(val_str)
        tipo = "A" if is_negative else "C"

        if f == int(f):
            valor_formatado = str(int(f))
        else:
            valor_formatado = f"{f:.2f}".replace(".", ",")

        return valor_formatado, tipo
    except ValueError:
        return val_str, "C"


def parse_nc_seq(nc_str):
    match = re.search(r'2026NC(\d+)', str(nc_str).strip())
    return match.group(1) if match else ""


# ─── LER CSV (SIMULANDO O FFILL DO PANDAS MANUALMENTE) ───────────────────────

def ler_csv_ffill(path):
    linhas = []
    last_valid = {}

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=",")

        for row in reader:
            # 1. Aplicar FFILL manual para cada coluna
            for key in row.keys():
                val = str(row.get(key, "")).strip()
                if val:
                    last_valid[key] = val
                else:
                    row[key] = last_valid.get(key, "")

            # 2. Extrai e trata os dados
            dt_emissao = row.get("Emissão - Dia", "")
            if not dt_emissao:
                continue

            valor_formatado, tipo_nota = clean_valor_e_tipo(row.get("Saldo - Moeda Origem (Item Informação)", ""))

            linhas.append({
                "dt_emissao": dt_emissao,
                "ecd_siafi": clean_int(row.get("Emitente - UG", "")),
                "esfera": clean_int(row.get("Esfera Orçamentária", "1")),
                "ptres": clean_int(row.get("PTRES", "")),
                "fonte": clean_fonte(row.get("Fonte Recursos Detalhada", "")),
                "natureza": clean_int(row.get("Natureza Despesa", "")),
                "pi": str(row.get("PI", "")).strip(),
                "rcd_siafi": clean_int(row.get("UG Responsável", "")),
                "nc_seq": parse_nc_seq(row.get("NC", "")),
                "obs": str(row.get("Doc - Observação", "")).strip(),
                "valor": valor_formatado,
                "tipo_nota": tipo_nota
            })
    return linhas


# ─── HELPERS: GENEXUS E DIGITAÇÃO HUMANA ─────────────────────────────────────

def gx_fill(page, field_id, value):
    """Preenche inputs injetando o valor direto via JS para driblar as máscaras."""
    if not value: return
    page.locator(f"#{field_id}").wait_for(state="visible", timeout=8000)
    page.evaluate("""([id, val]) => {
        var el = document.getElementById(id);
        if (!el) return;
        el.focus();
        el.value = val;
        if (typeof gx !== 'undefined') {
            gx.evt.onchange(el);
            gx.evt.onblur(el);
        } else {
            if(el.onchange) el.onchange();
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.dispatchEvent(new Event('blur', {bubbles: true}));
        }
    }""", [field_id, str(value)])


def gx_select(page, field_id, value):
    """Seleciona opções em Dropdowns (select) e avisa o Genexus."""
    if not value: return
    page.locator(f"#{field_id}").wait_for(state="visible", timeout=8000)
    page.select_option(f"#{field_id}", value=str(value))
    page.evaluate("""(id) => {
        var el = document.getElementById(id);
        if (!el) return;
        if (typeof gx !== 'undefined') {
            gx.evt.onchange(el);
        } else {
            el.dispatchEvent(new Event('change', {bubbles: true}));
        }
    }""", field_id)


def fill_data_humano(page, field_id, data_str):
    """Clica (seleciona tudo em azul), digita sobrescrevendo e dá TAB pra sair."""
    if not data_str: return
    page.locator(f"#{field_id}").click()
    apenas_numeros = data_str.replace("/", "").replace("-", "")
    page.keyboard.type(apenas_numeros, delay=50)
    page.keyboard.press("Tab")


# ─── MAIN ────────────────────────────────────────────────────────────────────

linhas = ler_csv_ffill(CSV_PATH)
print(f"📄 Linhas no CSV prontas para processar: {len(linhas)}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=800, channel="chrome")
    context = browser.new_context(ignore_https_errors=True)

    # Timeout geral de rede aumentado
    context.set_default_timeout(60000)
    page = context.new_page()

    # LOGIN
    print("⏳ Acessando o SILOMS...")
    page.goto("https://mac2.siloms.intraer/siloms_mac/servlet/aqs05019wse?E,S", wait_until="domcontentloaded")
    # Carrega as variáveis do arquivo .env
    load_dotenv()
    # Recupera os dados
    usuario = os.getenv("SISTEMA_USER")
    senha = os.getenv("SISTEMA_PASS")

    page.wait_for_selector("#username", state="visible")
    page.fill("#username", usuario)
    page.fill("#password", senha)
    page.check("#chkRemember")
    page.click("button[value='Entrar']")
    print("⏳ Entrando...")

    # NAVEGA PARA A TELA ALVO
    page.goto("https://mac2.siloms.intraer/siloms_mac/servlet/aqs06011w?S", wait_until="domcontentloaded")
    print("✅ Em aqs06011w")

    for idx, item in enumerate(linhas, start=1):
        print(f"\n[{idx}/{len(linhas)}] NC seq={item['nc_seq']} | Valor={item['valor']} | Tipo={item['tipo_nota']}")

        if "aqs06011w" not in page.url:
            page.goto("https://mac2.siloms.intraer/siloms_mac/servlet/aqs06011w?S", wait_until="domcontentloaded")

        page.wait_for_selector("input[name='BTNCADASTRARNC']", timeout=15000)
        page.click("input[name='BTNCADASTRARNC']")

        # Espera a tela de cadastro carregar
        page.wait_for_selector("#vDT_EMISSAO_CREDITO", state="visible", timeout=30000)

        # ── 1. DROPDOWNS ────────────────────────────────────────────────────
        gx_select(page, "vNR_ESFERA_CREDITO", item["esfera"])
        gx_select(page, "vTP_NOTA_CREDITO", item["tipo_nota"])

        # ── 2. DATA (digitação humana sobrescrevendo azul) ───────────────────
        data_final = item["dt_emissao"]
        if "-" in data_final and len(data_final) >= 10:
            partes = data_final.split(" ")[0].split("-")
            if len(partes[0]) == 4:
                data_final = f"{partes[2]}/{partes[1]}/{partes[0]}"

        fill_data_humano(page, "vDT_EMISSAO_CREDITO", data_final)

        # ── 3. CAMPOS DE TEXTO (Injeção JS rápida) ───────────────────────────
        gx_fill(page, "vECD_SIAFI", item["ecd_siafi"])
        gx_fill(page, "vNR_PTRES_CREDITO", item["ptres"])
        gx_fill(page, "vNR_FONTE_CREDITO", item["fonte"])
        gx_fill(page, "vCD_NATUREZA_CREDITO", item["natureza"])
        gx_fill(page, "vNR_PLANO_INTERNO_CREDITO", item["pi"])
        gx_fill(page, "vRCD_SIAFI", item["rcd_siafi"])
        gx_fill(page, "vFCD_SIAFI", "120633")  # fixo (GAP-SP)

        if item["nc_seq"]:
            gx_fill(page, "vSEQ", item["nc_seq"])

        gx_fill(page, "vTX_OBS_CREDITO", item["obs"])
        gx_fill(page, "vVL_TOTAL_CREDITO", item["valor"])

        print(f"   ✅ Campos inseridos!")

        # ⚠️ Deixei o input comentado pra rodar de vez, mas se quiser pausar pra olhar antes de salvar, é só descomentar!
        #input("👀 Confira a tela. ENTER para salvar...")

        # ── 4. SALVAR E VOLTAR ───────────────────────────────────────────────
        page.locator("#BTNENTER").wait_for(state="visible", timeout=10000)

        # Clica no botão verde de salvar e espera ele processar
        page.click("#BTNENTER")
        print("   ⏳ Salvando no Genexus...")
        page.wait_for_timeout(3000)

        # Volta pra lista pra pegar a próxima linha
        page.goto("https://mac2.siloms.intraer/siloms_mac/servlet/aqs06011w?S", wait_until="domcontentloaded")
        page.wait_for_selector("input[name='BTNCADASTRARNC']", state="visible", timeout=20000)
        print(f"   🔄 Pronto para próximo cadastro!")


    def processar_fila_conferencia(page):
        print("\n🚀 Iniciando Modo Instantâneo (Fila de Conferência)...")

        # Garante que começa na página certa uma única vez
        if not page.url.endswith("aqs06011w?S"):
            page.goto("https://mac2.siloms.intraer/siloms_mac/servlet/aqs06011w?S")

        while True:
            # --- PASSO 1: ATUALIZAR LISTA ---
            # Em vez de goto, clicamos no binóculos para atualizar a lista
            # Se não estiver em Registrada, ele seleciona, senão só clica
            status_dropdown = page.locator("#vST_NOTA_CREDITO")
            status_dropdown.wait_for(state="visible", timeout=5000)

            if status_dropdown.evaluate("node => node.value") != "A":
                gx_select(page, "vST_NOTA_CREDITO", "A")

            page.click("#IMAGE1", force=True)  # force=True pula verificações de visibilidade do Playwright

            # --- PASSO 2: BUSCAR ITEM ---
            btn_editar = page.locator("input[name^='vIMGEDIT']").first
            try:
                # Espera agressiva: apenas 2 segundos. Se não aparecer, acabou.
                btn_editar.wait_for(state="visible", timeout=2000)
            except:
                print("\n✅ FIM DA FILA: Tudo processado!")
                break

            # --- PASSO 3: ALOCAR (PROCESSO INTERNO) ---
            btn_editar.click()

            # Espera o campo de valor aparecer
            campo_valor = page.locator("#vVL_TOTAL_CREDITO")
            campo_valor.wait_for(state="visible", timeout=5000)

            valor = campo_valor.evaluate("node => node.value")
            gx_fill(page, "vGVL_CREDITO", valor)

            # Lupa/Indicadores
            page.click("#BTNPROMPTCONTACORRENTE")

            # Iframe: Espera apenas o que for necessário
            iframe = page.frame_locator("iframe").first
            try:
                # Tenta clicar no indicador link com timeout baixíssimo
                indicador = iframe.locator("span[id^='span_NR_DIGITO'] a").first
                indicador.wait_for(state="visible", timeout=1500)
                indicador.click()
            except:
                # Se falhar (não tem indicador), clica no "X" (fora do iframe)
                print("   ⚠️ Sem indicador. Pulando via 'X'...")
                page.locator("#gxp0_cls").click()

            # Preenchimento rápido
            gx_fill(page, "vDCD_PROJETO", "AU")
            ug_coord = page.locator("#vRCD_SIAFI").evaluate("node => node.value")
            gx_fill(page, "vDCD_SIAFI", ug_coord)

            # Confirmar Distribuição
            page.click("input[name='BTNINCLUIRCC']")

            # Checa se abriu popup de Novo Indicador
            iframe_novo = page.frame_locator("iframe").first
            btn_okay = iframe_novo.locator("#BTNENTER")
            if btn_okay.is_visible(timeout=800):
                page.once("dialog", lambda d: d.accept())
                btn_okay.click()

            # --- PASSO 4: CONFERIR (SEM SAIR DA PÁGINA) ---
            # Em vez de goto, clica no botão Voltar/Sair (IMAGE3) para voltar à lista
            page.locator("#IMAGE3").click()

            # Já marca o checkbox que deve estar na tela (primeira posição)
            check = page.locator("input[name='vMARCA_0001']").first
            check.wait_for(state="visible", timeout=5000)
            check.check()

            page.click("input[name='BTNCADASTRARNC4']")

            # Botão final de Aprovação
            btn_confere = page.locator("input[name='BTNCONFERE']")
            btn_confere.wait_for(state="visible", timeout=5000)
            page.once("dialog", lambda d: d.accept())
            btn_confere.click()

            # Volta para a lista para o próximo ciclo
            page.locator("#IMAGE3").wait_for(state="visible")
            page.locator("#IMAGE3").click()

            print(f"🚀 NC Processada!")


    # Executa
    processar_fila_conferencia(page)
