import pyautogui
import cv2
import numpy as np
import pytesseract
from pytesseract import Output
import time
import os
import re
import tkinter
from tkinter import messagebox
import pandas as pd
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Configuração inicial
root = tkinter.Tk()
root.withdraw()

debug = True
debug_dir = "debug"
if not os.path.exists(debug_dir):
    os.makedirs(debug_dir)

# Configuração dos logs
logs_dir = "logs"
if not os.path.exists(logs_dir):
    os.makedirs(logs_dir)

def limpar_arquivos_antigos(diretorio, prefixo, max_arquivos=10):
    """
    Mantém apenas os max_arquivos mais recentes com determinado prefixo em um diretório
    
    Args:
        diretorio: Diretório onde estão os arquivos
        prefixo: Prefixo dos arquivos a serem gerenciados
        max_arquivos: Número máximo de arquivos a manter
    """
    # Lista todos os arquivos com o prefixo especificado
    arquivos = [f for f in os.listdir(diretorio) if f.startswith(prefixo)]
    
    # Se houver mais arquivos que o limite
    if len(arquivos) > max_arquivos:
        # Ordena por data de modificação (mais antigo primeiro)
        arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)))
        
        # Remove os arquivos mais antigos
        for arquivo in arquivos[:-max_arquivos]:
            try:
                os.remove(os.path.join(diretorio, arquivo))
                print(f"Arquivo antigo removido: {arquivo}")
            except Exception as e:
                print(f"Erro ao remover arquivo {arquivo}: {e}")

# Gera nomes únicos para os arquivos de log desta execução
timestamp_execucao = datetime.now().strftime("%Y%m%d_%H%M%S")
nome_arquivo_log = f"{logs_dir}/log_{timestamp_execucao}.txt"
nome_arquivo_caminhos = f"{logs_dir}/caminhos_{timestamp_execucao}.txt"

# Conjunto para rastrear caminhos já registrados
caminhos_registrados = set()

# Limpa arquivos antigos no início da execução
limpar_arquivos_antigos(logs_dir, "log_", 10)
limpar_arquivos_antigos(logs_dir, "caminhos_", 10)

def registrar_log(mensagem, tipo="INFO"):
    """
    Registra uma mensagem no arquivo de log
    
    Args:
        mensagem: Mensagem a ser registrada
        tipo: Tipo da mensagem (INFO, ERRO, AVISO)
    """
    data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Formata a mensagem
    log_entry = f"[{data}] [{tipo}] {mensagem}\n"
    
    # Salva no arquivo
    with open(nome_arquivo_log, "a", encoding='utf-8') as f:
        f.write(log_entry)
    
    # Se for erro, também mostra no console
    if tipo == "ERRO":
        print(f"❌ {mensagem}")
    elif tipo == "AVISO":
        print(f"⚠️ {mensagem}")

def registrar_caminho(projeto, pasta_nivel, pasta_requisitos, dominio, pasta_use_case, sub_pasta=None, vf_nome=None, baixada=None, pasta_vazia=False):
    """
    Registra o caminho completo percorrido até uma pasta ou VF
    
    Args:
        projeto: Nome do projeto
        pasta_nivel: Pasta de maior nível (Work in Progress)
        pasta_requisitos: Pasta de requisitos funcionais
        dominio: Nome do domínio
        pasta_use_case: Nome do use case
        sub_pasta: Nome da sub-pasta (opcional)
        vf_nome: Nome da VF (opcional)
        baixada: Indica se a VF foi baixada com sucesso (opcional)
        pasta_vazia: Indica se é uma pasta vazia (opcional)
    """
    # Monta o caminho
    caminho = f"Projects\\{projeto}\\{pasta_nivel}\\{pasta_requisitos}\\{dominio}\\{pasta_use_case}"
    if sub_pasta:
        caminho += f"\\{sub_pasta}"
    if vf_nome:
        caminho += f"\\{vf_nome}"
        
    # Cria uma chave única para o caminho (sem o emoji)
    caminho_chave = caminho
    
    # Se este caminho já foi registrado e é uma VF que está sendo baixada,
    # só registra novamente se o status mudou de False para True
    if caminho_chave in caminhos_registrados:
        if not (vf_nome and vf_nome.split('_V')[0] in VFs and baixada):
            return
    
    # Adiciona o emoji apropriado
    if vf_nome:
        if vf_nome.split('_V')[0] in VFs:  # VF que deve ser baixada
            if baixada:
                caminho += " ✅"
            else:
                caminho += " 📄"  # Muda de ❌ para 📄 para indicar que ainda não foi baixada
        else:  # VF encontrada mas não está na lista para baixar
            caminho += " 📄"
    elif pasta_vazia:  # Se é uma pasta e está vazia
        caminho += " 📁"
    
    # Salva no arquivo
    with open(nome_arquivo_caminhos, "a", encoding='utf-8') as f:
        f.write(f"{caminho}\n")
    
    # Registra que este caminho já foi processado
    caminhos_registrados.add(caminho_chave)

def extrair_codigo_vf(nome_completo):
    """
    Extrai o código base da VF (ex: VF126 de VF126_V1_R6_P332BEV)
    """
    match = re.match(r'(VF\d+)', nome_completo)
    if match:
        return match.group(1)
    return nome_completo

def extrair_versao_vf(nome_completo):
    """
    Extrai a versão da VF (ex: VF126_V1 de VF126_V1_R6_P332BEV)
    """
    match = re.match(r'(VF\d+_V\d+)', nome_completo)
    if match:
        return match.group(1)
    return extrair_codigo_vf(nome_completo)

def criar_planilha_vfs(projetos):
    """
    Cria uma planilha Excel para rastrear as VFs encontradas
    
    Args:
        projetos: Lista de projetos que serão as colunas da planilha
    """
    # Colunas fixas iniciais
    colunas = ['Folder', 'VF'] + projetos
    
    # Criar DataFrame vazio
    df = pd.DataFrame(columns=colunas)
    
    # Nome do arquivo com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"Climate_VFs_{timestamp}.xlsx"
    
    # Salvar arquivo
    df.to_excel(nome_arquivo, index=False)
    print(f"✅ Planilha criada: {nome_arquivo}")
    return df, nome_arquivo

def adicionar_vf_planilha(df, nome_arquivo, folder, vf_info, projeto):
    """
    Adiciona ou atualiza uma VF na planilha
    """
    nome_completo = vf_info['texto_original']
    nome_base = extrair_versao_vf(nome_completo)  # Retorna VFxxx_Vx
    
    # Trabalha com uma cópia do DataFrame
    df_temp = df.copy()
    
    # Procura se já existe uma linha com esta VF nesta pasta
    linha_existente = None
    for idx in df_temp.index:
        if (df_temp.loc[idx, 'Folder'] == folder and 
            df_temp.loc[idx, 'VF'] == nome_base):
            linha_existente = idx
            break
    
    if linha_existente is not None:
        # Atualiza o nome completo no projeto correspondente
        df_temp.loc[linha_existente, projeto] = nome_completo
    else:
        # Cria nova linha
        nova_linha = pd.Series('', index=df_temp.columns)
        nova_linha['Folder'] = folder
        nova_linha['VF'] = nome_base
        nova_linha[projeto] = nome_completo
        df_temp = pd.concat([df_temp, pd.DataFrame([nova_linha])], ignore_index=True)
    
    # Ordena o DataFrame por pasta e nome da VF
    df_temp = df_temp.sort_values(['Folder', 'VF'])
    
    # Atualiza o DataFrame original
    df = df_temp.copy()
    
    # Salvar planilha após cada atualização
    df.to_excel(nome_arquivo, index=False)
    return df

def moveAndClick(image, clickType, offset_x=0, offset_y=0):
    """
    Move o mouse para uma imagem na tela e clica nela, com opção de offset.
    
    Args:
        image: Nome do arquivo de imagem a procurar
        clickType: Tipo de clique ('left', 'right', ou 'double')
        offset_x: Deslocamento em pixels no eixo X (positivo = direita, negativo = esquerda)
        offset_y: Deslocamento em pixels no eixo Y (positivo = baixo, negativo = cima)
    
    Returns:
        bool: True se a imagem foi encontrada e clicada, False caso contrário
    """
    # Captura uma captura de tela
    screenshot = pyautogui.screenshot()
    screenshot = np.array(screenshot)
    screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2GRAY)

    # Carrega a imagem de referência e a converte para escala de cinza
    template = cv2.imread(r"images/"+image, cv2.IMREAD_GRAYSCALE)

    # Usa correspondência de modelo para encontrar a posição
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # Define um limite de similaridade
    threshold = 0.7
    if max_val >= threshold:
        x, y = max_loc
        w, h = template.shape[::-1]
        center_x = x + w // 2 + offset_x  # Adiciona offset_x ao centro
        center_y = y + h // 2 + offset_y  # Adiciona offset_y ao centro
        
        if clickType == "left":
            pyautogui.leftClick(center_x, center_y, duration=0.5)
        elif clickType == "right":
            pyautogui.rightClick(center_x, center_y, duration=0.5)
        elif clickType == "double":
            pyautogui.doubleClick(center_x, center_y, duration=0.5)
        return True
    else:
        print(f"Imagem não encontrada com OpenCV. Confiança: {max_val:.2f}")
        return False

def mapear_pastas(icone_path, iniX, iniY, fimX, fimY):
    """
    Mapeia todas as pastas na interface, identificando seus nomes
    e coordenadas.
    
    Args:
        icone_path: Caminho para o arquivo de ícone
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca
    
    Returns:
        dict: Dicionário com as pastas mapeadas e suas coordenadas
    """
    print("🔍 Mapeando todas as pastas na interface...")
    #Mover o mouse para o centro da tela
    screen_width, screen_height = pyautogui.size()
    pyautogui.moveTo(screen_width/2, screen_height/2)
    time.sleep(1)
    # Verifica se o arquivo de referência do ícone existe
    if not os.path.exists(icone_path):
        mensagem = f"Arquivo de ícone '{icone_path}' não encontrado!"
        registrar_log(mensagem, "ERRO")
        return {}
    
    # Carrega a imagem do ícone e converte para escala de cinza
    icone_pasta = cv2.imread(icone_path, cv2.IMREAD_GRAYSCALE)
    if icone_pasta is None:
        mensagem = f"Erro ao carregar o ícone de pasta!"
        registrar_log(mensagem, "ERRO")
        return {}
    
    # Captura a tela
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    screenshot_cv = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
    
    # Identifica a região da árvore de pastas no DOORS
    altura, largura = screenshot_cv.shape[:2]
    
    # A área da árvore de pastas está no lado esquerdo
    inicio_x = int(largura * iniX)
    fim_x = int(largura * fimX)
    inicio_y = int(altura * iniY)
    fim_y = int(altura * fimY)
    
    # Recorta a região da árvore e converte para escala de cinza
    regiao_arvore = screenshot_cv[inicio_y:fim_y, inicio_x:fim_x]
    regiao_arvore_gray = cv2.cvtColor(regiao_arvore, cv2.COLOR_BGR2GRAY)
    
    # Encontra todas as ocorrências do ícone de pasta na árvore
    result = cv2.matchTemplate(regiao_arvore_gray, icone_pasta, cv2.TM_CCOEFF_NORMED)
    
    # Abaixa o limite para correspondência
    threshold = 0.65
    
    # Método de detecção de máximos locais
    kernel = np.ones((5, 5), np.uint8)
    dilated = cv2.dilate(result, kernel)
    matches = np.where((result >= threshold) & (result == dilated))
    pontos = list(zip(*matches[::-1]))
    
    # Melhor sistema de agrupamento
    pontos_filtrados = []
    icone_w, icone_h = icone_pasta.shape[:2]
    
    # Ordena os pontos por valor de correspondência
    pontos_com_score = [(pt[0], pt[1], result[pt[1], pt[0]]) for pt in pontos]
    pontos_com_score.sort(key=lambda x: x[2], reverse=True)
    
    # Agrupa os pontos por linha (coordenada Y)
    linhas = {}
    for x, y, score in pontos_com_score:
        linha_existente = None
        for linha_y in linhas.keys():
            if abs(y - linha_y) < icone_h // 2:
                linha_existente = linha_y
                break
        
        if linha_existente is None:
            linhas[y] = [(x, y, score)]
        else:
            linhas[linha_existente].append((x, y, score))
    
    # Para cada linha, ordena os pontos por coordenada X
    for linha_y, pontos_linha in linhas.items():
        pontos_linha.sort(key=lambda p: p[0])
        
        # Filtra pontos muito próximos na mesma linha
        x_anterior = -100
        for x, y, score in pontos_linha:
            if x - x_anterior >= icone_w * 0.8:
                pontos_filtrados.append((x, y))
                x_anterior = x
    
    print(f"Detectados {len(pontos_filtrados)} ícones de pasta após filtro")
    
    if not pontos_filtrados:
        print("❌ Nenhum ícone de pasta encontrado na tela!")
        return {}
    
    # Configuração do OCR
    config_ocr = r'--oem 1 --psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.'
    
    # Dicionário para armazenar os resultados
    pastas_mapeadas = {}
    
    # Imagem para visualização do mapeamento
    debug_regioes = regiao_arvore.copy()
    
    for idx, (x, y) in enumerate(pontos_filtrados):
        # Define a região de interesse à direita do ícone
        # Move a área de análise mais para a esquerda para capturar melhor o texto
        roi_y_start = max(0, y)
        roi_y_end = min(regiao_arvore.shape[0], y + icone_h)
        
        # Ajuste para começar mais próximo ao ícone (menos pixels à direita)
        roi_x_start = x + icone_w -1  # Inicia 1 pixels antes do final do ícone
        roi_x_end = min(regiao_arvore.shape[1], x + icone_w + 300)  # Limita a largura
        
        # Verifica se a região é válida
        if roi_x_start >= roi_x_end or roi_y_start >= roi_y_end:
            continue
        
        # Extrai a região de interesse
        roi = regiao_arvore[roi_y_start:roi_y_end, roi_x_start:roi_x_end]
        
        # Converte para escala de cinza
        roi_gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        
        # Inverte a imagem para texto branco em fundo preto (melhora OCR)
        roi_inv = cv2.bitwise_not(roi_gray)
        
        # Aplica OCR na região - usando a imagem invertida por padrão
        try:
            # Primeira tentativa com a imagem invertida e binarizada
            texto = pytesseract.image_to_string(roi_inv, config=config_ocr).strip()
                
            # Se ainda não encontrou, tenta com a imagem original
            if not texto:
                texto = pytesseract.image_to_string(roi_gray, config=config_ocr).strip()
            
            # Remove caracteres indesejados
            texto = re.sub(r'[^a-zA-Z0-9\-_.]', '', texto)
            
            # Se encontrou algum texto válido
            if texto:
                # Adiciona na imagem de debug
                cv2.putText(debug_regioes, f"{idx}:{texto}", (roi_x_start, roi_y_start - 2), 
                            cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 0), 1)
                
                # Marca região
                cv2.rectangle(debug_regioes, 
                            (roi_x_start, roi_y_start), 
                            (roi_x_end, roi_y_end), 
                            (0, 0, 255), 1)
                
                print(f"Pasta {idx}: '{texto}'")
                
                # Calcula a posição absoluta para clicar (ajustada para o novo ROI)
                click_x = inicio_x + x + icone_w + 20  # 20 pixels à direita do ícone
                click_y = inicio_y + y + icone_h // 2  # Centro vertical do ícone
                
                # Armazena no dicionário
                pastas_mapeadas[texto] = {
                    'x': click_x,
                    'y': click_y,
                    'texto_original': texto,
                    'icone_x': inicio_x + x,
                    'icone_y': inicio_y + y
                }
        except Exception as e:
            print(f"Erro ao processar região do ícone {idx}: {e}")
        
        # Salva as imagens de processamento para debug
        cv2.imwrite(f"{debug_dir}/roi_icone_{idx}_original.png", roi)
        cv2.imwrite(f"{debug_dir}/roi_icone_{idx}_inv.png", roi_inv)

    
    # Salva a imagem com as pastas mapeadas
    cv2.imwrite(f"{debug_dir}/pastas_mapeadas.png", debug_regioes)
    
    print(f"✅ Mapeamento concluído! {len(pastas_mapeadas)} pastas encontradas.")
    
    return pastas_mapeadas

def clicar_pasta(nome_pasta, mapa_pastas):
    """
    Clica em uma pasta específica usando o nome como referência.
    """
    print(f"🖱️ Buscando pasta '{nome_pasta}'")
    
    if mapa_pastas is None or len(mapa_pastas) == 0:
        mensagem = f"Não há mapeamento de pastas disponível para '{nome_pasta}'"
        registrar_log(mensagem, "ERRO")
        return False
    
    # Busca pela pasta no mapa
    correspondencia_exata = None
    correspondencia_parcial = None
    melhor_score = 0
    
    for texto, info in mapa_pastas.items():
        # Correspondência exata
        if texto == nome_pasta:
            correspondencia_exata = info
            break
        
        # Correspondência case-insensitive
        elif texto.lower() == nome_pasta.lower():
            if melhor_score < 90:
                correspondencia_parcial = info
                melhor_score = 90
        
        # Contém o nome da pasta como substring
        elif nome_pasta in texto:
            if melhor_score < 80:
                correspondencia_parcial = info
                melhor_score = 80
        
        # Nome da pasta contém o texto como substring (se texto for significativo)
        elif len(texto) > 3 and texto in nome_pasta:
            if melhor_score < 70:
                correspondencia_parcial = info
                melhor_score = 70
    
    # Usa correspondência exata se encontrou, ou a melhor correspondência parcial
    pasta_encontrada = correspondencia_exata or correspondencia_parcial
    
    if pasta_encontrada:
        pasta_texto = pasta_encontrada['texto_original']
        registrar_log(f"Pasta '{pasta_texto}' encontrada e clicada com sucesso", "INFO")
        
        # Clica na pasta
        click_x = pasta_encontrada['x']
        click_y = pasta_encontrada['y']
        
        screenshot = pyautogui.screenshot()
        screenshot_np = np.array(screenshot)
        screenshot_cv = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
        
        # Marca o local do clique
        cv2.circle(screenshot_cv, (click_x, click_y), 10, (0, 255, 0), -1)
        cv2.rectangle(
            screenshot_cv, 
            (pasta_encontrada['icone_x'], pasta_encontrada['icone_y']), 
            (pasta_encontrada['icone_x'] + 120, pasta_encontrada['icone_y'] + 20), 
            (0, 0, 255), 2
        )
        
        # Adiciona texto indicando a pasta
        cv2.putText(screenshot_cv, 
                   f"Clicando em: {pasta_texto}", 
                   (click_x + 20, click_y - 20), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                   
        if debug:
            cv2.imwrite(f"{debug_dir}/clique_pasta.png", screenshot_cv)
        
        print(f"🖱️ Clicando em ({click_x}, {click_y})")
        time.sleep(0.5)
        pyautogui.doubleClick(click_x, click_y, duration=0.5)
        return True
    else:
        mensagem = f"Pasta '{nome_pasta}' não encontrada no mapeamento"
        registrar_log(mensagem, "ERRO")
        return False

def esperarPor(image, timeout=10, iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95, imagem_interrupcao=None, 
             interrupcao_iniX=None, interrupcao_iniY=None, interrupcao_fimX=None, interrupcao_fimY=None):
    """
    Espera pela aparição de uma imagem em uma região específica da tela.
    Se uma imagem_interrupcao for fornecida e encontrada, retorna False.
    
    Args:
        image: Nome do arquivo de imagem a ser procurado na pasta 'images/'
        timeout: Tempo máximo de espera em segundos
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca para imagem principal
        imagem_interrupcao: Nome do arquivo de imagem que, se encontrado, interrompe a espera
        interrupcao_iniX, interrupcao_iniY, interrupcao_fimX, interrupcao_fimY: Coordenadas para busca da imagem de interrupção
    """
    start_time = time.time()
    
    # Carrega a imagem de referência
    template = cv2.imread(r"images/"+image, cv2.IMREAD_GRAYSCALE)
    if template is None:
        mensagem = f"Erro ao carregar imagem '{image}'. Verifique se existe em 'images/'"
        registrar_log(mensagem, "ERRO")
        return False
    
    # Carrega a imagem de interrupção se fornecida
    template_interrupcao = None
    if imagem_interrupcao:
        template_interrupcao = cv2.imread(r"images/"+imagem_interrupcao, cv2.IMREAD_GRAYSCALE)
        if template_interrupcao is None:
            mensagem = f"Erro ao carregar imagem de interrupção '{imagem_interrupcao}'"
            registrar_log(mensagem, "ERRO")
            return False
        
        # Se coordenadas específicas não foram fornecidas, usa as mesmas da imagem principal
        if interrupcao_iniX is None:
            interrupcao_iniX = iniX
        if interrupcao_iniY is None:
            interrupcao_iniY = iniY
        if interrupcao_fimX is None:
            interrupcao_fimX = fimX
        if interrupcao_fimY is None:
            interrupcao_fimY = fimY
    
    while time.time() - start_time < timeout:
        # Captura uma captura de tela completa
        screenshot = pyautogui.screenshot()
        screenshot_np = np.array(screenshot)
        
        # Calcula coordenadas da região de busca para imagem principal
        altura, largura = screenshot_np.shape[:2]
        inicio_x = int(largura * iniX)
        fim_x = int(largura * fimX)
        inicio_y = int(altura * iniY)
        fim_y = int(altura * fimY)
        
        # Recorta a região da tela para imagem principal
        regiao = screenshot_np[inicio_y:fim_y, inicio_x:fim_x]
        regiao_gray = cv2.cvtColor(regiao, cv2.COLOR_RGB2GRAY)
        
        if debug:
            # Salva apenas a região recortada usada na comparação
            cv2.imwrite(f"{debug_dir}/regiao_{image}.png", regiao_gray)
        
        # Verifica primeiro se a imagem de interrupção foi encontrada
        if template_interrupcao is not None:
            # Calcula coordenadas da região de busca para imagem de interrupção
            inicio_x_int = int(largura * interrupcao_iniX)
            fim_x_int = int(largura * interrupcao_fimX)
            inicio_y_int = int(altura * interrupcao_iniY)
            fim_y_int = int(altura * interrupcao_fimY)
            
            # Recorta a região para imagem de interrupção
            regiao_interrupcao = screenshot_np[inicio_y_int:fim_y_int, inicio_x_int:fim_x_int]
            regiao_interrupcao_gray = cv2.cvtColor(regiao_interrupcao, cv2.COLOR_RGB2GRAY)
            
            if debug:
                # Salva apenas a região recortada usada na comparação
                cv2.imwrite(f"{debug_dir}/regiao_{imagem_interrupcao}.png", regiao_interrupcao_gray)
            
            result_interrupcao = cv2.matchTemplate(regiao_interrupcao_gray, template_interrupcao, cv2.TM_CCOEFF_NORMED)
            min_val_int, max_val_int, min_loc_int, max_loc_int = cv2.minMaxLoc(result_interrupcao)
            if max_val_int >= 0.8:  # threshold
                mensagem = f"Imagem de interrupção '{imagem_interrupcao}' encontrada"
                registrar_log(mensagem, "AVISO")
                return False
        
        # Usa correspondência de modelo para encontrar a imagem principal
        result = cv2.matchTemplate(regiao_gray, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # Define um limite de similaridade
        threshold = 0.7
        if max_val >= threshold:
            time.sleep(1)
            return True
        
        time.sleep(1)
    
    mensagem = f"Timeout de {timeout} segundos: '{image}' não encontrado"
    registrar_log(mensagem, "AVISO")
    return False

def baixarVF(nome_VF):
    """
    Baixa uma VF e salva como arquivo Excel
    
    Args:
        nome_VF: Nome original da VF
        
    Returns:
        bool: True se o download foi bem sucedido, False caso contrário
    """
    # Trata o nome do arquivo para remover caracteres inválidos
    # Mantém apenas letras, números, underscores e hífens
    nome_arquivo = re.sub(r'[^\w\-]', '', nome_VF.replace('.', '_'))
    
    registrar_log(f"Iniciando download da VF: {nome_VF}", "INFO")
    
    #Organizar a VF
    if not moveAndClick("main.png", "right"):
        registrar_log(f"Falha ao clicar em 'main.png' para VF {nome_VF}", "ERRO")
        return False
    esperarPor("novo.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
    moveAndClick("novo.png", "left")
    esperarPor("barra.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_heading.png", "left")
    time.sleep(0.5)
    moveAndClick("inserir.png", "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_number.png", "left")
    time.sleep(0.5)
    moveAndClick("inserir.png", "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_level.png", "left")
    time.sleep(0.5)
    moveAndClick("inserir.png", "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_identifier.png", "left")
    time.sleep(0.5)
    moveAndClick("name.png", "left")
    time.sleep(0.5)
    pyautogui.write("RegID")
    time.sleep(0.5)
    moveAndClick("inserir.png", "left")
    time.sleep(0.5)
    moveAndClick("fechar.png", "left")
    esperarPor("coluna_object_heading.png", timeout=10, iniX=0.1, iniY=0.1, fimX=1, fimY=0.4)
    moveAndClick("coluna_object_heading.png", "right", offset_x=100)
    esperarPor("propriedades.png", timeout=10, iniX=0.3, iniY=0.05, fimX=1, fimY=0.5)
    moveAndClick("propriedades.png", "left")
    esperarPor("nome_main.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
    moveAndClick("nome_main.png", "left")
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.press('backspace')
    time.sleep(1)
    pyautogui.write("Main")
    time.sleep(1)
    moveAndClick("ok.png", "left")
    time.sleep(1)
    #Exportar a VF
    moveAndClick("arquivo.png", "left")
    time.sleep(1)
    moveAndClick("exportar.png", "left")
    time.sleep(1)
    moveAndClick("office.png", "left")
    time.sleep(1)
    moveAndClick("excel.png", "left")
    time.sleep(1)
    moveAndClick("barra_export.png", "left")
    time.sleep(1)
    moveAndClick("heading_and_text.png", "left")
    time.sleep(1)
    moveAndClick("check.png", "left")
    time.sleep(1)
    moveAndClick("exportar_excel.png", "left")
    time.sleep(10)
    #Salvar o Excel
    timeout = esperarPor("excel_icone.png", timeout=1200, iniX=0, iniY=0.5, fimX=1, fimY=0.98, imagem_interrupcao="fechar_erro.png", interrupcao_iniX=0, interrupcao_iniY=0.3, interrupcao_fimX=0.5, interrupcao_fimY=0.98)
    if not timeout:
        mensagem = f"Falha ao exportar VF {nome_VF}: Excel não encontrado ou erro ocorreu"
        registrar_log(mensagem, "ERRO")
        moveAndClick("fechar_erro.png", "left")
        time.sleep(2)
        moveAndClick("close_VF.png", "left")
        esperarPor("continuar_close_vf.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.8, fimY=0.95)
        moveAndClick("continuar_close_vf.png", "left")
        time.sleep(2)
        return False
    else:
        time.sleep(2)
        moveAndClick("excel_icone.png", "right")
        esperarPor("close_excel.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
        moveAndClick("close_excel.png", "left")
        esperarPor("save_excel.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
        moveAndClick("save_excel.png", "left")
        esperarPor("save_excel2.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
        pyautogui.write(nome_arquivo)
        time.sleep(1)
        moveAndClick("save_excel2.png", "left")
        time.sleep(5)
        moveAndClick("close_VF.png", "left")
        esperarPor("continuar_close_vf.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.8, fimY=0.95)
        moveAndClick("continuar_close_vf.png", "left")
        time.sleep(2)
        registrar_log(f"Download da VF {nome_VF} concluído com sucesso", "INFO")
        return True

def get_pasta_nivel(nome_pasta):
    """
    Determina o nível hierárquico de uma pasta.
    Work in Progress é o nível mais alto.
    Para as demais, o número determina o nível primário e a letra o nível secundário.
    """
    # Normaliza o nome da pasta para comparação
    nome_normalizado = nome_pasta.lower().strip()
    
    # Verifica se contém "old" - deve ser ignorada
    if 'old' in nome_normalizado:
        return (float('-inf'), 0)  # Retorna -infinito para garantir que nunca seja escolhida
    
    # "Work in Progress" tem prioridade máxima
    # Verifica várias possíveis variações do texto
    if any(termo in nome_normalizado for termo in ['work in progress', 'work_in_progress', 'workinprogress', 'work-in-progress']):
        return (float('inf'), 0)  # Retorna infinito para garantir que seja sempre o maior
    
    # Tenta extrair o número e a letra (ex: "1A", "2B")
    match = re.match(r'.*?(\d+)([A-Za-z])', nome_normalizado)
    if match:
        numero = int(match.group(1))
        letra = ord(match.group(2).upper()) - ord('A')  # Converte letra para número (A=0, B=1, etc)
        return (numero, letra)
    
    return (-1, -1)  # Retorna nível mínimo para pastas que não seguem o padrão

def encontrar_pasta_maior_nivel(mapa_pastas):
    """
    Encontra a pasta de maior nível hierárquico no mapa de pastas,
    ignorando pastas que contenham a palavra 'old'.
    """
    maior_nivel = (float('-inf'), 0)
    pasta_escolhida = None
    
    print("\nAnalisando níveis das pastas:")
    for nome_pasta in mapa_pastas.keys():
        nivel_atual = get_pasta_nivel(nome_pasta)
        print(f"Pasta: {nome_pasta} -> Nível: {nivel_atual}")
        
        # Ignora pastas com 'old'
        if 'old' in nome_pasta.lower():
            print(f"Ignorando pasta com 'old': {nome_pasta}")
            continue
            
        if nivel_atual > maior_nivel:
            maior_nivel = nivel_atual
            pasta_escolhida = nome_pasta
    
    if pasta_escolhida:
        print(f"\nPasta escolhida: {pasta_escolhida} (Nível: {maior_nivel})")
    else:
        print("\n❌ Nenhuma pasta válida encontrada")
    
    return pasta_escolhida

def encontrar_pasta_requisitos(mapa_pastas):
    """
    Encontra a pasta de requisitos funcionais entre as pastas mapeadas.
    Aceita várias variações do nome e ignora pastas que contenham 'old'.
    """
    termos_requisitos = [
        'functional requirements',
        'functional_requirements',
        'functionalrequirements',
        'functional-requirements',
        'functional req',
        'func requirements',
        'func req'
    ]
    
    for nome_pasta, info in mapa_pastas.items():
        # Remove 'folder' do nome para todas as comparações
        nome_limpo = nome_pasta.lower().replace('folder', '').strip()
        
        # Ignora pastas que contenham 'old' (após remover 'folder')
        if 'old' in nome_limpo:
            print(f"Ignorando pasta com 'old': {nome_pasta}")
            continue
            
        if any(termo in nome_limpo for termo in termos_requisitos):
            print(f"✅ Pasta de requisitos encontrada: {nome_pasta}")
            return nome_pasta
    
    print("❌ Pasta de requisitos funcionais não encontrada")
    return None

def voltar_nivel(nivel):
    # Obtém as dimensões da tela
    screen_width, screen_height = pyautogui.size()
    # Clica no meio da tela
    pyautogui.click(screen_width/2, screen_height/2)
    time.sleep(0.5)
    pyautogui.hotkey('shift', 'tab')
    time.sleep(1)
    for i in range(nivel):
        pyautogui.hotkey('shift', 'left')  # Usa hotkey para pressionar shift + seta esquerda
    time.sleep(1)

projetos = ["226MCA","291","521MCA","521MY21","598","363","341","281","3580","2651"]#["139EL","250MY24","250MY26","312MCA","332BEV","334MCA","356MCA","356MHEV","520MY24","637BEV","637MCA","846","965","ALFAMCA","ARM20","ARM23","LP3","M240","M240MY26-BEV","MASAHMCA","MASAHMY26","332TR"]
dominios = ["Climate", "Comfort Climate"]
use_cases = []
VFs = ["VF126"]

# Inicializar planilha de rastreamento
df_vfs, nome_arquivo_vfs = criar_planilha_vfs(projetos)

time.sleep(5)

#Clicar no botão de projetos
if not moveAndClick("projects.png", "left"):
    print("❌ Parando")
    messagebox.showerror("Timeout", "O reconhecimento de imagem falhou")        
    exit()

#Mapear as pastas
if esperarPor("pasta.png"):
    pastas_projetos = mapear_pastas(icone_path="images/pasta.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
else:
    print("❌ Parando")
    messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
    exit()

for projeto in projetos:
    # Clica no projeto
    clicar_pasta(projeto, pastas_projetos)
    time.sleep(1)
    
    # Mapeia as subpastas do projeto
    if esperarPor("pasta.png"):
        pastas_niveis = mapear_pastas(icone_path="images/pasta.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        
        # Encontra e clica na pasta de maior nível
        pasta_maior_nivel = encontrar_pasta_maior_nivel(pastas_niveis)
        if pasta_maior_nivel:
            print(f"Selecionando pasta de maior nível: {pasta_maior_nivel}")
            clicar_pasta(pasta_maior_nivel, pastas_niveis)
        else:
            registrar_log(f"Nenhuma pasta válida encontrada no projeto {projeto}", "ERRO")
            messagebox.showerror("Erro", "Nenhuma pasta válida encontrada no projeto")
            exit()
    else:
        registrar_log("Falha ao mapear pastas de nível", "ERRO")
        messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
        exit()

    # Procura e clica em Functional Requirements
    if esperarPor("pasta_amarela.png"):
        time.sleep(1)
        pastas_requerimentos = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        
        pasta_requisitos = encontrar_pasta_requisitos(pastas_requerimentos)
        if pasta_requisitos:
            if not clicar_pasta(pasta_requisitos, pastas_requerimentos):
                registrar_log(f"Erro ao clicar na pasta de requisitos em {projeto}", "ERRO")
                messagebox.showerror("Erro", "Erro ao clicar na pasta de requisitos funcionais")
                exit()
            time.sleep(2)
        else:
            registrar_log(f"Pasta de requisitos funcionais não encontrada em {projeto}", "ERRO")
            messagebox.showerror("Erro", "Pasta de requisitos funcionais não encontrada")
            
            exit()
    else:
        registrar_log("Falha ao mapear pasta de requisitos", "ERRO")
        messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
        exit()

    if esperarPor("pasta_amarela.png"):
        pastas_dominios = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        
        # Verifica se algum domínio da lista foi encontrado nas pastas
        dominio_encontrado = None
        for nome_pasta in pastas_dominios:
            # Remove 'folder' do nome para comparação
            nome_limpo = nome_pasta.replace(" ", "").lower().replace('folder', '').strip()
            for dominio in dominios:
                if dominio.replace(" ", "").lower() == nome_limpo:
                    dominio_encontrado = nome_pasta
                    break
            if dominio_encontrado:
                break
        
        if dominio_encontrado:
            registrar_log(f"Domínio encontrado: {dominio_encontrado}", "INFO")
            clicar_pasta(dominio_encontrado, pastas_dominios)
            time.sleep(1)
            if esperarPor("pasta_amarela.png"):
                pastas_use_cases = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                print(pastas_use_cases)
                
                # Se a lista use_cases não estiver vazia, procura apenas os use cases listados
                if use_cases:
                    use_cases_encontrados = []
                    for nome_pasta in pastas_use_cases:
                        nome_normalizado = nome_pasta.replace(" ", "").lower()
                        for use_case in use_cases:
                            if use_case.replace(" ", "").lower() == nome_normalizado:
                                use_cases_encontrados.append(nome_pasta)
                                break
                    
                    # Atualiza a lista de use cases para processar apenas os encontrados
                    if use_cases_encontrados:
                        registrar_log(f"Use cases encontrados: {', '.join(use_cases_encontrados)}", "INFO")
                        pastas_para_processar = use_cases_encontrados
                    else:
                        registrar_log(f"Nenhum dos use cases especificados foi encontrado", "AVISO")
                        continue
                else:
                    # Se use_cases estiver vazia, processa todos os use cases encontrados
                    pastas_para_processar = list(pastas_use_cases.keys())
                
                # Processa os use cases
                for use_case in pastas_para_processar:
                    clicar_pasta(use_case, pastas_use_cases)
                    time.sleep(1)

                    sub_pastas = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                    vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)

                    # Se não encontrou nem subpastas nem VFs, é uma pasta vazia
                    if not sub_pastas and not vf_nomes:
                        registrar_caminho(
                            projeto=projeto,
                            pasta_nivel=pasta_maior_nivel,
                            pasta_requisitos=pasta_requisitos,
                            dominio=dominio_encontrado,
                            pasta_use_case=use_case,
                            pasta_vazia=True
                        )

                    # Registra VFs encontradas
                    for vf_nome, vf_info in vf_nomes.items():
                        df_vfs = adicionar_vf_planilha(
                            df_vfs,
                            nome_arquivo_vfs,
                            folder=use_case,
                            vf_info=vf_info,
                            projeto=projeto
                        )
                        
                        # Verifica se é uma VF que deve ser baixada
                        baixar = False
                        for vf_esperado in VFs:
                            if vf_nome.lower().startswith(vf_esperado.lower()):
                                baixar = True
                                break
                        
                        # Registra o caminho
                        registrar_caminho(
                            projeto=projeto,
                            pasta_nivel=pasta_maior_nivel,
                            pasta_requisitos=pasta_requisitos,
                            dominio=dominio_encontrado,
                            pasta_use_case=use_case,
                            vf_nome=vf_nome,
                            baixada=False if baixar else None  # None para VFs que não precisam ser baixadas
                        )
                        
                        if baixar:
                            clicar_pasta(vf_nome, vf_nomes)
                            time.sleep(1)
                            moveAndClick("abrir_somente_leitura.png", "left")
                            esperarPor("main.png", timeout=20, iniX=0.1, iniY=0.1, fimX=0.9, fimY=0.5)
                            sucesso = baixarVF(vf_nome)
                            # Atualiza o registro com o status do download
                            registrar_caminho(
                                projeto=projeto,
                                pasta_nivel=pasta_maior_nivel,
                                pasta_requisitos=pasta_requisitos,
                                dominio=dominio_encontrado,
                                pasta_use_case=use_case,
                                vf_nome=vf_nome,
                                baixada=sucesso
                            )

                    if sub_pastas != {}:
                        for sub_pasta in sub_pastas:
                            clicar_pasta(sub_pasta, sub_pastas)
                            vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                            
                            # Se não encontrou VFs na subpasta, é uma pasta vazia
                            if not vf_nomes:
                                registrar_caminho(
                                    projeto=projeto,
                                    pasta_nivel=pasta_maior_nivel,
                                    pasta_requisitos=pasta_requisitos,
                                    dominio=dominio_encontrado,
                                    pasta_use_case=use_case,
                                    sub_pasta=sub_pasta,
                                    pasta_vazia=True
                                )
                            
                            # Registra VFs encontradas na sub-pasta
                            for vf_nome, vf_info in vf_nomes.items():
                                df_vfs = adicionar_vf_planilha(
                                    df_vfs,
                                    nome_arquivo_vfs,
                                    folder=f"{use_case}/{sub_pasta}",
                                    vf_info=vf_info,
                                    projeto=projeto
                                )
                                
                                # Verifica se é uma VF que deve ser baixada
                                baixar = False
                                for vf_esperado in VFs:
                                    if vf_nome.lower().startswith(vf_esperado.lower()):
                                        baixar = True
                                        break
                                
                                # Registra o caminho
                                registrar_caminho(
                                    projeto=projeto,
                                    pasta_nivel=pasta_maior_nivel,
                                    pasta_requisitos=pasta_requisitos,
                                    dominio=dominio_encontrado,
                                    pasta_use_case=use_case,
                                    sub_pasta=sub_pasta,
                                    vf_nome=vf_nome,
                                    baixada=False if baixar else None  # None para VFs que não precisam ser baixadas
                                )
                                
                                if baixar:
                                    clicar_pasta(vf_nome, vf_nomes)
                                    time.sleep(1)
                                    moveAndClick("abrir_somente_leitura.png", "left")
                                    esperarPor("main.png", timeout=20, iniX=0.1, iniY=0.1, fimX=0.9, fimY=0.5)
                                    sucesso = baixarVF(vf_nome)
                                    # Atualiza o registro com o status do download
                                    registrar_caminho(
                                        projeto=projeto,
                                        pasta_nivel=pasta_maior_nivel,
                                        pasta_requisitos=pasta_requisitos,
                                        dominio=dominio_encontrado,
                                        pasta_use_case=use_case,
                                        sub_pasta=sub_pasta,
                                        vf_nome=vf_nome,
                                        baixada=sucesso
                                    )

                            time.sleep(1)
                            voltar_nivel(1)

                        time.sleep(1)   
                        voltar_nivel(2)
                    else:
                        time.sleep(1)
                        voltar_nivel(1)
        else:
            time.sleep(1)
            voltar_nivel(1)
    else:               
        print("❌ Parando")
        messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
        exit()
    
    voltar_nivel(7)
    moveAndClick("projects.png", "left")
    time.sleep(1)

print(f"\n✅ Processo concluído! A planilha foi salva em: {nome_arquivo_vfs}")
messagebox.showinfo("Concluído", f"Processo finalizado!\nA planilha foi salva em:\n{nome_arquivo_vfs}")

