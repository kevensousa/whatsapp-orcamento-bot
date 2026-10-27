import os
import glob
import re
import time
import random
import urllib.parse
import configparser
import winreg
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")
MSG_FILE = os.path.join(BASE_DIR, "mensagem.txt")
EXCEL_FILE = os.path.join(BASE_DIR, "contatos.xlsx")
CSV_FILE = os.path.join(BASE_DIR, "contatos.csv")
RELATORIO_FILE = os.path.join(BASE_DIR, "relatorio_envios.xlsx")
PASTA_IMAGENS = os.path.join(BASE_DIR, "imagens")

def carregar_configuracoes():
    """Carrega as configurações do arquivo config.ini"""
    config = configparser.ConfigParser()
    padrao = {
        "navegador": "firefox",
        "delay_min": 8,
        "delay_max": 15,
        "imagem_padrao": "exemplo.jpeg"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            config.read(CONFIG_FILE, encoding="utf-8")
            if "CONFIGURACOES" in config:
                padrao["navegador"] = config.get("CONFIGURACOES", "navegador", fallback="firefox").strip().lower()
                padrao["delay_min"] = config.getint("CONFIGURACOES", "delay_minimo", fallback=8)
                padrao["delay_max"] = config.getint("CONFIGURACOES", "delay_maximo", fallback=15)
                padrao["imagem_padrao"] = config.get("CONFIGURACOES", "imagem_padrao", fallback="exemplo.jpeg").strip()
        except Exception as e:
            print(f"⚠️ Aviso ao ler config.ini ({e}), usando valores padrão.")
            
    return padrao

def detectar_navegador_windows():
    """Detecta o navegador padrão associado no Registro do Windows"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice") as key:
            prog_id, _ = winreg.QueryValueEx(key, "ProgId")
            prog_id = prog_id.lower()
            if "firefox" in prog_id:
                return "firefox"
            elif "chrome" in prog_id:
                return "chrome"
            elif "edge" in prog_id or "msedge" in prog_id:
                return "edge"
    except Exception:
        pass
    return None

def definir_navegador(opcao_config):
    """Define qual navegador iniciar"""
    opcao = opcao_config.lower().strip()
    if opcao in ["firefox", "chrome", "edge"]:
        return opcao
        
    padrao_win = detectar_navegador_windows()
    if padrao_win:
        return padrao_win
        
    caminhos_firefox = [
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Mozilla Firefox\firefox.exe")
    ]
    if any(os.path.exists(c) for c in caminhos_firefox):
        return "firefox"
        
    return "chrome"

def limpar_numero(numero_raw):
    """Limpa e formata o número de telefone no padrão internacional (DDI + DDD + Número)"""
    if not numero_raw:
        return None
    numero = re.sub(r"\D", "", str(numero_raw).split(".")[0])
    
    if not numero:
        return None
        
    if len(numero) in (10, 11):
        numero = "55" + numero
    elif len(numero) in (12, 13) and numero.startswith("55"):
        pass
    elif len(numero) < 10:
        return None
        
    return numero

def resolver_caminho_imagem(imagem_nome):
    """Localiza o arquivo de imagem mesmo com diferença de extensão (.png/.jpeg/.jpg)"""
    if not imagem_nome or pd.isna(imagem_nome):
        imagens_disponiveis = [f for f in glob.glob(os.path.join(PASTA_IMAGENS, "*")) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp'))]
        if imagens_disponiveis:
            return os.path.abspath(imagens_disponiveis[0])
        return None

    imagem_str = str(imagem_nome).strip()
    if not imagem_str:
        return None

    # 1. Correspondência direta
    candidatos = [
        imagem_str,
        os.path.join(PASTA_IMAGENS, imagem_str),
        os.path.join(BASE_DIR, imagem_str)
    ]
    for c in candidatos:
        if os.path.isfile(c):
            return os.path.abspath(c)

    # 2. Correspondência flexível de extensão
    nome_sem_ext = os.path.splitext(imagem_str)[0]
    padroes_busca = [
        os.path.join(PASTA_IMAGENS, f"{nome_sem_ext}.*"),
        os.path.join(BASE_DIR, f"{nome_sem_ext}.*"),
        f"{nome_sem_ext}.*"
    ]
    for padrao in padroes_busca:
        arquivos = glob.glob(padrao)
        for arq in arquivos:
            if arq.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                return os.path.abspath(arq)

    # 3. Fallback: primeira foto disponível na pasta 'imagens'
    imagens_disponiveis = [f for f in glob.glob(os.path.join(PASTA_IMAGENS, "*")) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp'))]
    if imagens_disponiveis:
        return os.path.abspath(imagens_disponiveis[0])

    print(f"⚠️ Imagem '{imagem_str}' não encontrada na pasta 'imagens'.")
    return None

def carregar_mensagem():
    """Carrega o texto padrão de mensagem"""
    if os.path.exists(MSG_FILE):
        with open(MSG_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    msg_example = os.path.join(BASE_DIR, "mensagem.example.txt")
    if os.path.exists(msg_example):
        with open(msg_example, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "Olá! Gostaria de solicitar um orçamento para prestação de serviços. Poderia me enviar os valores e disponibilidade?"

def carregar_contatos():
    """Carrega a planilha de contatos (Excel ou CSV)"""
    df = None
    if os.path.exists(EXCEL_FILE):
        print(f"📄 Carregando contatos de: {EXCEL_FILE}")
        df = pd.read_excel(EXCEL_FILE)
    elif os.path.exists(CSV_FILE):
        print(f"📄 Carregando contatos de: {CSV_FILE}")
        try:
            df = pd.read_csv(CSV_FILE, sep=";")
            if len(df.columns) == 1:
                df = pd.read_csv(CSV_FILE, sep=",")
        except Exception:
            df = pd.read_csv(CSV_FILE, sep=",")
    else:
        # Fallback para o arquivo de exemplo
        example_xlsx = os.path.join(BASE_DIR, "contatos.example.xlsx")
        if os.path.exists(example_xlsx):
            print(f"📄 Carregando contatos de modelo: {example_xlsx}")
            df = pd.read_excel(example_xlsx)
        else:
            print("❌ Nenhuma planilha encontrada! Crie 'contatos.xlsx' ou 'contatos.csv'.")
            return None

    # Normalizar nomes de colunas
    colunas_map = {}
    for col in df.columns:
        col_norm = col.strip().lower()
        if col_norm in ["telefone", "celular", "numero", "número", "contato", "phone", "whatsapp"]:
            colunas_map[col] = "telefone"
        elif col_norm in ["nome", "prestador", "empresa", "fornecedor", "name"]:
            colunas_map[col] = "nome"
        elif col_norm in ["imagem", "foto", "anexo", "image", "photo", "arquivo"]:
            colunas_map[col] = "imagem"
            
    df = df.rename(columns=colunas_map)
    
    if "telefone" not in df.columns:
        print("⚠️ Coluna de telefone não identificada pelo nome. Usando a primeira coluna como número.")
        df = df.rename(columns={df.columns[0]: "telefone"})
        
    return df

def fechar_popups_iniciais(page):
    """Fecha avisos iniciais como 'Novidades do WhatsApp Web'"""
    try:
        botoes = page.query_selector_all(
            "button:has-text('Continuar'), div[role='button']:has-text('Continuar'), button:has-text('Agora não')"
        )
        for btn in botoes:
            if btn.is_visible():
                btn.click()
                time.sleep(0.5)
    except Exception:
        pass

def enviar_imagem_whatsapp(page, caminho_imagem):
    """Envia uma imagem no chat aberto como FOTO oficial (não como figurinha)"""
    try:
        time.sleep(1.5)

        # 1. Clica no botão de Anexar ('+') no rodapé
        btn_anexo = page.query_selector("footer button[aria-label*='Anexar'], button[aria-label*='Anexar']")
        if not btn_anexo:
            print("⚠️ Botão Anexar não localizado no rodapé.")
            return False, "Botão Anexar não encontrado"

        btn_anexo.click()
        time.sleep(1)

        # 2. Clica especificamente em 'Fotos e vídeos' usando expect_file_chooser
        btn_fotos = page.query_selector("button[aria-label*='Fotos e vídeos'], button:has-text('Fotos e vídeos')")
        if btn_fotos:
            with page.expect_file_chooser() as fc_info:
                btn_fotos.click()
            file_chooser = fc_info.value
            file_chooser.set_files(caminho_imagem)
            print(f"🖼️ Foto '{os.path.basename(caminho_imagem)}' carregada em modo IMAGEM REAL. Aguardando prévia...")
        else:
            # Fallback de segurança se o botão de menu não for achado
            file_input = page.query_selector("input[type='file'][accept*='image'], input[type='file']")
            if not file_input:
                print("⚠️ Campo de anexo não encontrado no WhatsApp Web.")
                return False, "Input de arquivo não localizado"
            file_input.set_input_files(caminho_imagem)
            print(f"🖼️ Foto carregada via input direto.")

        # 3. Aguarda o botão verde de envio da prévia de foto (Enviar 1 item selecionado)
        btn_send_media = page.wait_for_selector(
            "div[role='button'][aria-label*='Enviar'], div[aria-label*='Enviar'], [data-testid='send']",
            timeout=12000
        )

        if btn_send_media and btn_send_media.is_visible():
            time.sleep(0.8)
            btn_send_media.click()
            print("✅ Clique no botão de envio da foto realizado!")
        else:
            time.sleep(0.8)
            page.keyboard.press("Enter")
            print("✅ Foto enviada via Enter!")

        # Aguarda upload e envio no chat
        time.sleep(4)
        print(f"✅ 2ª mensagem (Foto '{os.path.basename(caminho_imagem)}') enviada com sucesso no formato de IMAGEM!")
        return True, "Imagem enviada"

    except Exception as e:
        print(f"❌ Erro ao enviar imagem: {e}")
        return False, str(e)

def enviar_mensagens():
    configs = carregar_configuracoes()
    navegador = definir_navegador(configs["navegador"])
    delay_min = configs["delay_min"]
    delay_max = configs["delay_max"]
    img_padrao_config = configs["imagem_padrao"]

    msg_template = carregar_mensagem()
    df = carregar_contatos()
    
    if df is None or df.empty:
        print("❌ Lista de contatos vazia ou não encontrada. Encerrando.")
        return

    print(f"\n📋 Total de linhas na planilha: {len(df)}")
    print(f"🧭 Navegador selecionado: {navegador.upper()}")
    print("--------------------------------------------------")
    print("Mensagem modelo (1ª Mensagem):\n")
    print(msg_template)
    print("--------------------------------------------------\n")

    historico = []

    with sync_playwright() as p:
        if navegador == "firefox":
            print("🦊 Abrindo o Mozilla Firefox...")
            perfil_dir = os.path.join(BASE_DIR, "perfil_firefox")
            browser = p.firefox.launch_persistent_context(
                user_data_dir=perfil_dir,
                headless=False,
                viewport=None
            )
        elif navegador == "edge":
            print("🌐 Abrindo o Microsoft Edge...")
            perfil_dir = os.path.join(BASE_DIR, "perfil_edge")
            browser = p.chromium.launch_persistent_context(
                user_data_dir=perfil_dir,
                channel="msedge",
                headless=False,
                args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
                viewport=None
            )
        else:
            print("🌐 Abrindo o Google Chrome...")
            perfil_dir = os.path.join(BASE_DIR, "perfil_chrome")
            browser = p.chromium.launch_persistent_context(
                user_data_dir=perfil_dir,
                headless=False,
                args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
                viewport=None
            )

        page = browser.pages[0] if browser.pages else browser.new_page()
        print("⏳ Acessando WhatsApp Web...")
        page.goto("https://web.whatsapp.com")

        print("📱 Aguardando conexão ou leitura do QR Code...")
        try:
            page.wait_for_selector("#pane-side, [data-icon='chat'], [aria-label='Painel de conversas'], [role='textbox']", timeout=120000)
            fechar_popups_iniciais(page)
            print("✅ WhatsApp Web conectado com sucesso!")
        except Exception:
            print("⚠️ Tempo esgotado para login. Verifique se escaneou o QR Code.")
            browser.close()
            return

        time.sleep(2)

        total = len(df)
        for idx, row in df.iterrows():
            numero_original = row.get("telefone", "")
            nome = str(row.get("nome", "")).strip() if "nome" in df.columns and pd.notna(row.get("nome")) else ""
            img_contato = row.get("imagem", "") if "imagem" in df.columns and pd.notna(row.get("imagem")) else ""
            
            img_alvo = img_contato if img_contato else img_padrao_config
            caminho_imagem = resolver_caminho_imagem(img_alvo)

            numero = limpar_numero(numero_original)
            hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            if not numero:
                print(f"[{idx+1}/{total}] ⚠️ Número inválido ou vazio: '{numero_original}'. Pulando...")
                historico.append({
                    "Nome": nome,
                    "Telefone": numero_original,
                    "Telefone_Formatado": "",
                    "Status_Texto": "NÚMERO INVÁLIDO",
                    "Status_Imagem": "NÃO ENVIADO",
                    "Horário": hora_atual,
                    "Detalhes": "Formato incorreto ou menos de 10 dígitos"
                })
                continue

            # Personaliza a mensagem se houver a tag {nome}
            if "{nome}" in msg_template:
                msg_final = msg_template.replace("{nome}", nome)
            else:
                msg_final = msg_template

            msg_encoded = urllib.parse.quote(msg_final)
            url_chat = f"https://web.whatsapp.com/send?phone={numero}&text={msg_encoded}"

            print(f"\n[{idx+1}/{total}] 🚀 Enviando para {nome or 'Contato'} ({numero})...")
            
            try:
                page.goto(url_chat, wait_until="domcontentloaded")
                
                enviado_texto = False
                tempo_limite = 25
                inicio = time.time()

                while time.time() - inicio < tempo_limite:
                    # Verifica se o número não está no WhatsApp
                    popup_erro = page.query_selector("div[data-animate-modal-popup='true'], [data-testid='popup-contents'], div[role='dialog']")
                    if popup_erro:
                        texto_popup = popup_erro.inner_text().lower()
                        if any(termo in texto_popup for termo in ["não está no whatsapp", "inválido", "invalid", "not on whatsapp", "url é inválido", "url is invalid"]):
                            print(f"❌ O número {numero} não tem WhatsApp ou é inválido.")
                            btn_ok = page.query_selector("button:has-text('OK'), div[role='button']:has-text('OK'), button[aria-label='OK']")
                            if btn_ok:
                                try:
                                    btn_ok.click()
                                except Exception:
                                    pass
                            historico.append({
                                "Nome": nome,
                                "Telefone": numero_original,
                                "Telefone_Formatado": numero,
                                "Status_Texto": "NÃO POSSUI WHATSAPP",
                                "Status_Imagem": "NÃO ENVIADO",
                                "Horário": hora_atual,
                                "Detalhes": "Aviso do WhatsApp de número não registrado"
                            })
                            enviado_texto = False
                            break

                    # Verifica botão de envio do texto
                    btn_send = page.query_selector("span[data-icon='send'], button[aria-label*='Enviar'], button[data-tab='11'], [data-testid='send']")
                    if btn_send:
                        time.sleep(1)
                        btn_send.click()
                        time.sleep(2)
                        print(f"✅ 1ª mensagem (Texto) enviada para {numero}!")
                        enviado_texto = True
                        break

                    # Fallback com tecla Enter no campo de texto
                    caixa_texto = page.query_selector("footer div[contenteditable='true']")
                    if caixa_texto and caixa_texto.is_visible():
                        time.sleep(0.5)
                        caixa_texto.press("Enter")
                        time.sleep(2)
                        print(f"✅ 1ª mensagem (Texto) enviada via Enter para {numero}!")
                        enviado_texto = True
                        break

                    time.sleep(1)

                # Se o texto foi enviado com sucesso e há imagem configurada:
                status_img_final = "SEM IMAGEM CONFIGURADA"
                if enviado_texto:
                    if caminho_imagem:
                        print(f"📷 Enviando 2ª mensagem (foto: {os.path.basename(caminho_imagem)}) para {numero}...")
                        sucesso_img, msg_img = enviar_imagem_whatsapp(page, caminho_imagem)
                        status_img_final = "ENVIADA" if sucesso_img else f"FALHA ({msg_img})"
                    
                    historico.append({
                        "Nome": nome,
                        "Telefone": numero_original,
                        "Telefone_Formatado": numero,
                        "Status_Texto": "ENVIADO",
                        "Status_Imagem": status_img_final,
                        "Horário": hora_atual,
                        "Detalhes": f"Texto OK | Foto: {status_img_final}"
                    })
                elif not any(h["Telefone_Formatado"] == numero for h in historico if h["Horário"] == hora_atual):
                    print(f"⚠️ Tempo esgotado ao tentar enviar texto para {numero}.")
                    historico.append({
                        "Nome": nome,
                        "Telefone": numero_original,
                        "Telefone_Formatado": numero,
                        "Status_Texto": "FALHA / TIMEOUT",
                        "Status_Imagem": "NÃO ENVIADO",
                        "Horário": hora_atual,
                        "Detalhes": "Não foi possível carregar a conversa a tempo"
                    })

            except Exception as e:
                print(f"❌ Erro ao enviar para {numero}: {e}")
                historico.append({
                    "Nome": nome,
                    "Telefone": numero_original,
                    "Telefone_Formatado": numero,
                    "Status_Texto": "ERRO",
                    "Status_Imagem": "ERRO",
                    "Horário": hora_atual,
                    "Detalhes": str(e)
                })

            # Intervalo anti-spam entre contatos
            if idx < total - 1:
                delay = random.randint(delay_min, delay_max)
                print(f"⏳ Aguardando {delay} segundos antes do próximo contato...")
                time.sleep(delay)

        # Salva o relatório atualizado
        df_historico = pd.DataFrame(historico)
        df_historico.to_excel(RELATORIO_FILE, index=False)
        print("\n==================================================")
        print(f"🎉 Processo concluído! Relatório salvo em:\n{RELATORIO_FILE}")
        print("==================================================")

        time.sleep(3)
        browser.close()

if __name__ == "__main__":
    enviar_mensagens()
