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

debug = False
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
                print(f"Old file removed: {arquivo}")
            except Exception as e:
                print(f"Error removing file {arquivo}: {e}")

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
    Records a message in the log file
    
    Args:
        mensagem: Message to be recorded
        tipo: Message type (INFO, ERROR, WARNING)
    """
    data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Format the message
    log_entry = f"[{data}] [{tipo}] {mensagem}\n"
    
    # Save to file
    with open(nome_arquivo_log, "a", encoding='utf-8') as f:
        f.write(log_entry)
    
    # If it's an error, also show in console
    if tipo == "ERROR":
        print(f"❌ {mensagem}")
    elif tipo == "WARNING":
        print(f"⚠️ {mensagem}")

def registrar_caminho(projeto, pasta_nivel, pasta_requisitos, dominio, pasta_use_case, sub_pasta=None, vf_nome=None, baixada=None, pasta_vazia=False, vfs_list=None):
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
        vfs_list: Lista de VFs que devem ser baixadas (opcional)
    """
    # Se não foi fornecida uma lista de VFs, usa uma lista vazia
    if vfs_list is None:
        vfs_list = []
    
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
        if not (vf_nome and vf_nome.split('_V')[0] in vfs_list and baixada):
            return
    
    # Adiciona o emoji apropriado
    if vf_nome:
        if vf_nome.split('_V')[0] in vfs_list:  # VF que deve ser baixada
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
    print(f"✅ Spreadsheet created: {nome_arquivo}")
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

def moveAndClick(image, clickType, offset_x=0, offset_y=0, iniX=0, iniY=0, fimX=1, fimY=1):
    """
    Move o mouse para uma imagem na tela e clica nela, com opção de offset e região de busca.
    Aceita uma única imagem ou uma lista de imagens para buscar.
    
    Args:
        image: Nome do arquivo de imagem a procurar ou lista de nomes de imagens
        clickType: Tipo de clique ('left', 'right', ou 'double')
        offset_x: Deslocamento em pixels no eixo X (positivo = direita, negativo = esquerda)
        offset_y: Deslocamento em pixels no eixo Y (positivo = baixo, negativo = cima)
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca (0 a 1)
    
    Returns:
        bool: True se alguma imagem foi encontrada e clicada, False caso contrário
    """
    # Converte uma imagem única em uma lista para processamento uniforme
    if isinstance(image, str):
        images = [image]
    else:
        images = image
        
    # Captura uma captura de tela
    screenshot = pyautogui.screenshot()
    screenshot = np.array(screenshot)
    screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2GRAY)

    # Calcula as coordenadas absolutas da região de busca
    altura, largura = screenshot.shape[:2]
    inicio_x = int(largura * iniX)
    fim_x = int(largura * fimX)
    inicio_y = int(altura * iniY)
    fim_y = int(altura * fimY)
    
    # Recorta a região de busca
    regiao_busca = screenshot[inicio_y:fim_y, inicio_x:fim_x]
    
    # Define um limite de similaridade
    threshold = 0.7
    
    # Armazena o melhor resultado entre todas as imagens
    best_match = None
    best_score = 0
    best_image = None
    best_template = None
    
    # Tenta encontrar cada imagem na região
    for img in images:
        # Carrega a imagem de referência e a converte para escala de cinza
        template = cv2.imread(r"images/"+img, cv2.IMREAD_GRAYSCALE)
        if template is None:
            print(f"❌ Imagem '{img}' não encontrada!")
            continue

        # Usa correspondência de modelo para encontrar a posição
        result = cv2.matchTemplate(regiao_busca, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # Se esta imagem tem uma correspondência melhor que as anteriores
        if max_val >= threshold and max_val > best_score:
            best_score = max_val
            best_match = max_loc
            best_image = img
            best_template = template

    # Se encontrou alguma correspondência boa
    if best_match is not None:
        # Converte as coordenadas relativas à região para coordenadas absolutas da tela
        x, y = best_match
        w, h = best_template.shape[::-1]
        center_x = inicio_x + x + w // 2 + offset_x
        center_y = inicio_y + y + h // 2 + offset_y
        
        print(f"✅ Imagem '{best_image}' encontrada com confiança: {best_score:.2f}")
        
        if clickType == "left":
            pyautogui.leftClick(center_x, center_y, duration=0.5)
        elif clickType == "right":
            pyautogui.rightClick(center_x, center_y, duration=0.5)
        elif clickType == "double":
            pyautogui.doubleClick(center_x, center_y, duration=0.5)
        return True
    else:
        imagens_str = ", ".join(images)
        print(f"❌ Nenhuma das imagens [{imagens_str}] foi encontrada com confiança suficiente")
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
        registrar_log(mensagem, "ERROR")
        return {}
    
    # Carrega a imagem do ícone e converte para escala de cinza
    icone_pasta = cv2.imread(icone_path, cv2.IMREAD_GRAYSCALE)
    if icone_pasta is None:
        mensagem = f"Erro ao carregar o ícone de pasta!"
        registrar_log(mensagem, "ERROR")
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
        registrar_log(mensagem, "ERROR")
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
        registrar_log(f"Folder '{pasta_texto}' found and clicked successfully", "INFO")
        
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
        mensagem = f"Folder '{nome_pasta}' not found in the mapping"
        registrar_log(mensagem, "ERROR")
        return False

def esperarPor(image, timeout=30, iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95, imagem_interrupcao=None, 
             interrupcao_iniX=None, interrupcao_iniY=None, interrupcao_fimX=None, interrupcao_fimY=None):
    """
    Espera pela aparição de uma imagem em uma região específica da tela.
    Aceita uma única imagem ou uma lista de imagens para buscar.
    Se uma imagem_interrupcao for fornecida e encontrada, retorna False.
    
    Args:
        image: Nome do arquivo de imagem a ser procurado ou lista de nomes de imagens
        timeout: Tempo máximo de espera em segundos
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca para imagem principal
        imagem_interrupcao: Nome do arquivo de imagem ou lista de imagens que, se encontradas, interrompem a espera
        interrupcao_iniX, interrupcao_iniY, interrupcao_fimX, interrupcao_fimY: Coordenadas para busca da imagem de interrupção
        
    Returns:
        bool: True se alguma imagem foi encontrada, False caso contrário ou se houve interrupção
              A imagem encontrada é registrada no log
    """
    start_time = time.time()
    
    # Converte imagens únicas em listas para processamento uniforme
    if isinstance(image, str):
        images = [image]
    else:
        images = image
    
    # Processa imagens de interrupção
    interrupcao_images = []
    if imagem_interrupcao:
        if isinstance(imagem_interrupcao, str):
            interrupcao_images = [imagem_interrupcao]
        else:
            interrupcao_images = imagem_interrupcao
    
    # Carrega todas as imagens principais de uma vez
    templates = []
    for img in images:
        template = cv2.imread(r"images/"+img, cv2.IMREAD_GRAYSCALE)
        if template is None:
            mensagem = f"Image '{img}' loading error. Check if it exists in 'images/'"
            registrar_log(mensagem, "ERROR")
            continue
        templates.append((img, template))
    
    if not templates:
        registrar_log("None of the images could be loaded", "ERROR")
        return False
    
    # Carrega todas as imagens de interrupção
    interrupcao_templates = []
    for img in interrupcao_images:
        template = cv2.imread(r"images/"+img, cv2.IMREAD_GRAYSCALE)
        if template is None:
            mensagem = f"Erro ao carregar imagem de interrupção '{img}'"
            registrar_log(mensagem, "ERROR")
            continue
        interrupcao_templates.append((img, template))
        
    # Se coordenadas específicas não foram fornecidas, usa as mesmas da imagem principal
    if interrupcao_iniX is None:
        interrupcao_iniX = iniX
    if interrupcao_iniY is None:
        interrupcao_iniY = iniY
    if interrupcao_fimX is None:
        interrupcao_fimX = fimX
    if interrupcao_fimY is None:
        interrupcao_fimY = fimY
    
    # Define um limite de similaridade
    threshold = 0.7
    
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
        
        # Verifica primeiro se alguma imagem de interrupção foi encontrada
        if interrupcao_templates:
            # Calcula coordenadas da região de busca para imagem de interrupção
            inicio_x_int = int(largura * interrupcao_iniX)
            fim_x_int = int(largura * interrupcao_fimX)
            inicio_y_int = int(altura * interrupcao_iniY)
            fim_y_int = int(altura * interrupcao_fimY)
            
            # Recorta a região para imagem de interrupção
            regiao_interrupcao = screenshot_np[inicio_y_int:fim_y_int, inicio_x_int:fim_x_int]
            regiao_interrupcao_gray = cv2.cvtColor(regiao_interrupcao, cv2.COLOR_RGB2GRAY)
            
            # Verifica cada imagem de interrupção
            for nome_img, template in interrupcao_templates:
                if debug:
                    # Salva a região recortada usada na comparação
                    cv2.imwrite(f"{debug_dir}/regiao_{nome_img}.png", regiao_interrupcao_gray)
                
                result_interrupcao = cv2.matchTemplate(regiao_interrupcao_gray, template, cv2.TM_CCOEFF_NORMED)
                min_val_int, max_val_int, min_loc_int, max_loc_int = cv2.minMaxLoc(result_interrupcao)
                
                if max_val_int >= 0.8:  # threshold
                    mensagem = f"Interruption image '{nome_img}' found"
                    registrar_log(mensagem, "WARNING")
                    return False
        
        # Procura cada imagem principal
        for nome_img, template in templates:
            if debug:
                # Salva a região recortada usada na comparação
                cv2.imwrite(f"{debug_dir}/regiao_{nome_img}.png", regiao_gray)
            
            # Usa correspondência de modelo para encontrar a imagem principal
            result = cv2.matchTemplate(regiao_gray, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= threshold:
                mensagem = f"Image '{nome_img}' found with confidence: {max_val:.2f}"
                registrar_log(mensagem, "INFO")
                time.sleep(1)
                return True
        
        time.sleep(1)
    
    imagens_str = ", ".join([img for img, _ in templates])
    registrar_log(f"Timeout of {timeout} seconds: None of the images [{imagens_str}] was found", "WARNING")
    return False

def encontrar_coordenadas_y_main(iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.6):
    """
    Procura pela imagem main.png na tela e retorna suas coordenadas Y mais baixa (min)
    e mais alta (max) como valores relativos (0 a 1) em relação à altura da tela.
    
    Args:
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca (0 a 1)
        
    Returns:
        tuple: (y_min, y_max) - coordenadas Y mínima e máxima relativas da imagem main.png
               ou (0.1, 0.4) por padrão se a imagem não for encontrada
    """
    # Carrega a imagem de referência
    template = cv2.imread(r"images/main.png", cv2.IMREAD_GRAYSCALE)
    if template is None:
        mensagem = f"Erro ao carregar imagem 'main.png'. Verifique se existe em 'images/'"
        registrar_log(mensagem, "ERROR")
        return (0.1, 0.4)  # Valores padrão
    
    # Captura uma screenshot da tela
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    
    # Converte para escala de cinza
    screenshot_gray = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2GRAY)
    
    # Calcula as coordenadas absolutas da região de busca
    altura, largura = screenshot_gray.shape[:2]
    inicio_x = int(largura * iniX)
    fim_x = int(largura * fimX)
    inicio_y = int(altura * iniY)
    fim_y = int(altura * fimY)
    
    # Recorta a região de busca
    regiao_busca = screenshot_gray[inicio_y:fim_y, inicio_x:fim_x]
    
    # Usa correspondência de modelo para encontrar todas as ocorrências da imagem
    result = cv2.matchTemplate(regiao_busca, template, cv2.TM_CCOEFF_NORMED)
    
    # Define um limite de similaridade
    threshold = 0.7
    
    # Encontra todas as ocorrências acima do threshold
    locations = np.where(result >= threshold)
    
    if len(locations[0]) > 0:
        # Encontra o Y mínimo e máximo em coordenadas relativas à região de busca
        y_min_regiao = int(np.min(locations[0]))
        y_max_regiao = int(np.max(locations[0]) + template.shape[0])
        
        # Converte para coordenadas absolutas na tela
        y_min_abs = inicio_y + y_min_regiao
        y_max_abs = inicio_y + y_max_regiao
        
        # Converte para valores relativos à altura total da tela
        y_min_rel = y_min_abs / altura
        y_max_rel = y_max_abs / altura
        
        return (y_min_rel, y_max_rel)
    
    # Se não encontrou, retorna valores padrão
    registrar_log("Could not find Y coordinates of the main.png image", "WARNING")
    return (0.1, 0.4)

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
    
    registrar_log(f"Starting download of VF: {nome_VF}", "INFO")

    if esperarPor("maximizar_vf.png", timeout=5, iniX=0.1, iniY=0.05, fimX=0.7, fimY=0.5):
        moveAndClick("maximizar_vf.png", "left")
        time.sleep(0.5)
    
    esperarPor("main.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.4)
    
    # Encontra as coordenadas Y da imagem main.png
    y_min, y_max = encontrar_coordenadas_y_main()
    
    if esperarPor("separador_coluna.png", timeout=5, iniX=0.11, iniY=y_min, fimX=0.3, fimY=y_max):
        moveAndClick("separador_coluna.png", "right", offset_x=-50)
        time.sleep(0.5)
        moveAndClick(["remover.png", "remover_en.png"], "left")
        time.sleep(0.5)

    
    #Organizar a VF
    if not moveAndClick("main.png", "right"):
        moveAndClick("close_VF.png", "left")
        esperarPor(["continuar_close_vf.png", "continuar_close_vf_en.png"], timeout=10, iniX=0.05, iniY=0.05, fimX=0.8, fimY=0.95)
        moveAndClick(["continuar_close_vf.png", "continuar_close_vf_en.png"], "left")
        time.sleep(2)
        registrar_log(f"Failed to click on 'main.png' for VF {nome_VF}", "ERROR")
        return False
    esperarPor(["novo.png", "novo_en.png"], timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
    moveAndClick(["novo.png", "novo_en.png"], "left")
    esperarPor("barra.png", timeout=10, iniX=0.05, iniY=0.05, fimX=0.95, fimY=0.95)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_text.png", "left")
    time.sleep(0.5)
    moveAndClick("name.png", "left")
    time.sleep(0.5)
    pyautogui.write("Main")
    time.sleep(0.5)
    moveAndClick(["inserir.png", "inserir_en.png"], "left")
    time.sleep(0.5)
    moveAndClick("name_main.png", "left")
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.press('backspace')
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_heading.png", "left")
    time.sleep(0.5)
    moveAndClick(["inserir.png", "inserir_en.png"], "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_number.png", "left")
    time.sleep(0.5)
    moveAndClick(["inserir.png", "inserir_en.png"], "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_level.png", "left")
    time.sleep(0.5)
    moveAndClick(["inserir.png", "inserir_en.png"], "left")
    time.sleep(0.5)
    moveAndClick("barra.png", "left")
    time.sleep(0.5)
    moveAndClick("object_identifier.png", "left")
    time.sleep(0.5)
    moveAndClick("name.png", "left")
    time.sleep(0.5)
    pyautogui.write("RegID")
    time.sleep(0.5)
    moveAndClick(["inserir.png", "inserir_en.png"], "left")
    time.sleep(0.5)
    moveAndClick(["fechar.png", "fechar_en.png"], "left")
    esperarPor("main_text.png", timeout=10, iniX=0.1, iniY=0.1, fimX=1, fimY=0.4)
    moveAndClick("main_text.png", "right", offset_x=200)
    esperarPor(["remover.png", "remover_en.png"], timeout=10, iniX=0.3, iniY=0.05, fimX=1, fimY=0.5)
    moveAndClick(["remover.png", "remover_en.png"], "left")
    time.sleep(1)
    #Exportar a VF
    moveAndClick(["arquivo.png", "arquivo_en.png"], "left")
    esperarPor(["exportar.png", "exportar_en.png"], timeout= 30, iniX=0, iniY=0, fimX=0.6, fimY=0.5)
    moveAndClick(["exportar.png", "exportar_en.png"], "left")
    esperarPor(["planilha_export.png", "planilha_export_en.png"], timeout= 30, iniX=0, iniY=0, fimX=0.6, fimY=0.5)
    moveAndClick(["planilha_export.png", "planilha_export_en.png"], "left")
    esperarPor(["procurar_export.png", "procurar_export_en.png"], timeout= 30, iniX=0.3, iniY=0.50, fimX=0.6, fimY=0.80)
    moveAndClick(["procurar_export.png", "procurar_export_en.png"], "left")
    esperarPor("desktop_export.png", timeout= 30, iniX=0.4, iniY=0.2, fimX=0.9, fimY=0.80)
    moveAndClick("desktop_export.png", "left")
    time.sleep(0.5)
    moveAndClick(["abrir_export.png", "abrir_export_en.png"], "left")
    esperarPor(["exportar_csv.png", "exportar_csv_en.png"], timeout= 30, iniX=0.3, iniY=0.50, fimX=0.6, fimY=0.80)
    moveAndClick(["exportar_csv.png", "exportar_csv_en.png"], "left")
    #Checando se tem repetido
    if esperarPor(["confirmar_sobrescrever.png", "confirmar_sobrescrever_en.png"], timeout= 5, iniX=0.25, iniY=0.3, fimX=0.7, fimY=0.7):
        moveAndClick(["confirmar_sobrescrever.png", "confirmar_sobrescrever_en.png"], "left")
        time.sleep(0.5)
    #Esperando o export acabar
    exportando = True
    while exportando:
        exportando = esperarPor("doors_icon.png", timeout= 5, iniX=0.3, iniY=0.20, fimX=0.6, fimY=0.4)
    #Fechando VF
    moveAndClick("close_VF.png", "left")
    esperarPor(["continuar_close_vf.png", "continuar_close_vf_en.png"], timeout=10, iniX=0.05, iniY=0.05, fimX=0.8, fimY=0.95)
    moveAndClick(["continuar_close_vf.png", "continuar_close_vf_en.png"], "left")
    time.sleep(2)
    registrar_log(f"Download of VF {nome_VF} completed successfully", "INFO")
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
        time.sleep(0.5)
        pyautogui.hotkey('shift', 'left')  # Usa hotkey para pressionar shift + seta esquerda
    time.sleep(0.5)

def encontrar_posicao_xy(image, iniX=0, iniY=0, fimX=1, fimY=1):
    """
    Procura por uma imagem na tela e retorna suas coordenadas x e y
    como percentual da largura e altura da tela.
    Aceita uma única imagem ou uma lista de imagens para buscar.
    
    Args:
        image: Nome do arquivo de imagem a procurar ou lista de nomes de imagens
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca
        
    Returns:
        tuple: (x_percentual, y_percentual) - Coordenadas como percentual da tela
               ou (None, None) se a imagem não for encontrada
    """
    # Converte uma imagem única em uma lista para processamento uniforme
    if isinstance(image, str):
        images = [image]
    else:
        images = image
    
    # Captura uma captura de tela completa
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    
    # Calcula coordenadas da região de busca
    altura, largura = screenshot_np.shape[:2]
    inicio_x = int(largura * iniX)
    fim_x = int(largura * fimX)
    inicio_y = int(altura * iniY)
    fim_y = int(altura * fimY)
    
    # Recorta a região da tela
    regiao = screenshot_np[inicio_y:fim_y, inicio_x:fim_x]
    
    # Define um limite de similaridade
    threshold = 0.7
    
    # Armazena o melhor resultado entre todas as imagens
    best_score = 0
    best_loc = None
    best_image = None
    
    # Tenta encontrar cada imagem na região
    for img in images:
        # Carrega a imagem de referência e a converte para escala de cinza
        template = cv2.imread(r"images/"+img, cv2.IMREAD_GRAYSCALE)
        if template is None:
            mensagem = f"Erro ao carregar imagem '{img}'. Verifique se existe em 'images/'"
            registrar_log(mensagem, "ERROR")
            continue
        
        # Converte a região de busca para escala de cinza
        regiao_gray = cv2.cvtColor(regiao, cv2.COLOR_RGB2GRAY)
        
        # Usa correspondência de modelo para encontrar a imagem
        result = cv2.matchTemplate(regiao_gray, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # Se esta imagem tem uma correspondência melhor que as anteriores
        if max_val >= threshold and max_val > best_score:
            best_score = max_val
            best_loc = max_loc
            best_image = img
    
    # Se encontrou alguma correspondência boa
    if best_loc is not None:
        # Calcula a posição x e y absoluta
        x_absoluto = inicio_x + best_loc[0]
        y_absoluto = inicio_y + best_loc[1]
        
        # Converte para percentual da tela
        x_percentual = x_absoluto / largura
        y_percentual = y_absoluto / altura
        
        registrar_log(f"Image '{best_image}' found with confidence: {best_score:.2f}", "INFO")
        return x_percentual, y_percentual
    
    imagens_str = ", ".join(images)
    registrar_log(f"None of the images [{imagens_str}] was found", "WARNING")
    return None, None

def procura_projeto(nome):
    
    esperarPor(["ferramentas.png", "ferramentas_en.png"], timeout=30, iniX=0.01, iniY= 0.05, fimX=0.7, fimY=0.4)
    moveAndClick(["ferramentas.png", "ferramentas_en.png"], "left")
    esperarPor(["localizar.png", "localizar_en.png"], timeout=30, iniX=0.01, iniY= 0.05, fimX=0.7, fimY=0.4)
    moveAndClick(["localizar.png", "localizar_en.png"], "left")
    esperarPor("check_localizar.png", timeout=30, iniX=0.3, iniY=0.2, fimX=7, fimY=0.8)
    pyautogui.write(nome)
    time.sleep(1)
    moveAndClick("check_localizar.png", "left")
    time.sleep(0.5)
    pyautogui.press("enter")
    if esperarPor("pasta.png",timeout=10, iniX=0.3, iniY=0.4, fimX=0.7, fimY=0.8):
        moveAndClick(["pasta.png", "pasta_en.png"], "double", iniX=0.3, iniY=0.4, fimX=0.7, fimY=0.8)
        time.sleep(2)
        moveAndClick(["fechar_localizar.png", "fechar_localizar_en.png"], "left")
        time.sleep(1)
        return True
    else:
        moveAndClick(["fechar_localizar.png", "fechar_localizar_en.png"], "left")
        time.sleep(1)
        return False

def filtrar_codigos_por_regiao(caminho_planilha, regiao):
    """
    Reads an Excel spreadsheet and filters codes by region and development phase
    
    Args:
        caminho_planilha: Path to the Excel file
        regiao: Region to filter (e.g., 'EMEA', 'NAFTA', etc.)
        
    Returns:
        list: List of "Old Code" values that match the region and have 
              "Development Phase" different from 'obsolete' or 'inactive'
              Empty values or '-' in "Old Code" are ignored
    """
    try:
        # Read the Excel file
        df = pd.read_excel(caminho_planilha)
        
        # Log the columns found in the file
        registrar_log(f"Columns found in the spreadsheet: {', '.join(df.columns)}", "INFO")
        
        # Check if required columns exist
        required_columns = ["Region", "Old Code", "Development Phase"]
        for col in required_columns:
            if col not in df.columns:
                registrar_log(f"Required column '{col}' not found in the spreadsheet", "ERROR")
                return []
        
        # Filter by region
        df_filtered = df[df["Region"] == regiao]
        
        if len(df_filtered) == 0:
            registrar_log(f"No entries found for region: {regiao}", "WARNING")
            return []
            
        # Filter out obsolete and inactive development phases
        df_filtered = df_filtered[~df_filtered["Development Phase"].str.lower().isin(["obsolete", "inactive"])]
        
        if len(df_filtered) == 0:
            registrar_log(f"No active entries found for region: {regiao}", "WARNING")
            return []
            
        # Filter out empty values or '-' in "Old Code"
        df_filtered = df_filtered[
            (~df_filtered["Old Code"].isna()) &  # Remove NaN values
            (df_filtered["Old Code"] != "") &    # Remove empty strings
            (df_filtered["Old Code"] != "-")     # Remove dash character
        ]
        
        # Log if any rows were filtered out due to empty or '-' values
        if len(df_filtered) < len(df[df["Region"] == regiao]):
            registrar_log(f"Filtered out {len(df[df['Region'] == regiao]) - len(df_filtered)} rows with empty or '-' values in 'Old Code'", "INFO")
        
        if len(df_filtered) == 0:
            registrar_log(f"No valid codes found for region: {regiao} after filtering", "WARNING")
            return []
            
        # Extract the "Old Code" values into a list
        codigos = df_filtered["Old Code"].tolist()
        
        # Log the number of codes found
        registrar_log(f"Found {len(codigos)} active codes for region: {regiao}", "INFO")
        
        return codigos
        
    except Exception as e:
        registrar_log(f"Error processing spreadsheet: {str(e)}", "ERROR")
        return []

def main_logic(projetos, dominios, use_cases, VFs):
    """
    Função principal do macro que executa a lógica de busca e download das VFs
    
    Args:
        projetos: Lista de códigos de projetos para pesquisar
        dominios: Lista de domínios para filtrar
        use_cases: Lista de casos de uso para filtrar
        VFs: Lista de VFs para baixar
    """
    # Inicializar planilha de rastreamento
    df_vfs, nome_arquivo_vfs = criar_planilha_vfs(projetos)

    time.sleep(5)

    # Clicando no botão projetos
    if not moveAndClick("projects.png", "left"):
        print("❌ Parando")
        messagebox.showerror("Timeout", "Image recognition failed")        
        return

    for projeto in projetos:

        if not procura_projeto(projeto):
            registrar_log(f"Project {projeto} not found. Continuing to the next project.", "WARNING")
            continue

        voltar = 7
        
        
        # Mapeia as subpastas do projeto
        if esperarPor("pasta.png"):
            pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
            # Usa valores padrão se não encontrou a imagem
            if pos_x is None:
                pos_x, pos_y = 0.3, 0.1  # Valores padrão
                registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
            
            pastas_niveis = mapear_pastas(icone_path="images/pasta.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
            
            # Encontra e clica na pasta de maior nível
            pasta_maior_nivel = encontrar_pasta_maior_nivel(pastas_niveis)
            if pasta_maior_nivel:
                print(f"Selecionando pasta de maior nível: {pasta_maior_nivel}")
                clicar_pasta(pasta_maior_nivel, pastas_niveis)
            else:
                registrar_log(f"No valid folder found in project {projeto}", "ERROR")
                messagebox.showerror("Error", "No valid folder found in project")
                return
        else:
            registrar_log("Failed to map level folders", "ERROR")
            messagebox.showerror("Timeout", "Image recognition failed")
            return

        # Procura e clica em Functional Requirements
        if esperarPor("pasta_amarela.png", timeout= 30):

            pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
            # Usa valores padrão se não encontrou a imagem
            if pos_x is None:
                pos_x, pos_y = 0.3, 0.1  # Valores padrão
                registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
                
            pastas_requerimentos = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
            
            pasta_requisitos = encontrar_pasta_requisitos(pastas_requerimentos)
            if pasta_requisitos:
                if not clicar_pasta(pasta_requisitos, pastas_requerimentos):
                    registrar_log(f"Error clicking on requirements folder in {projeto}", "ERROR")
                    messagebox.showerror("Error", "Error clicking on requirements folder")
                    return
                
            else:
                registrar_log(f"Functional requirements folder not found in {projeto}", "ERROR")
                messagebox.showerror("Error", "Functional requirements folder not found")
                return
        else:
            registrar_log("Failed to map requirements folder", "ERROR")
            messagebox.showerror("Timeout", "Image recognition failed")
            return

        if esperarPor("pasta_amarela.png"):
            pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
            # Usa valores padrão se não encontrou a imagem
            if pos_x is None:
                pos_x, pos_y = 0.3, 0.1  # Valores padrão
                registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
                
            pastas_dominios = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
            
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
                registrar_log(f"Domain found: {dominio_encontrado}", "INFO")
                clicar_pasta(dominio_encontrado, pastas_dominios)
                
                if esperarPor("pasta_amarela.png"):
                    pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
                    # Usa valores padrão se não encontrou a imagem
                    if pos_x is None:
                        pos_x, pos_y = 0.3, 0.1  # Valores padrão
                        registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
                        
                    pastas_use_cases = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
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
                            registrar_log(f"Use cases found: {', '.join(use_cases_encontrados)}", "INFO")
                            pastas_para_processar = use_cases_encontrados
                        else:
                            registrar_log(f"None of the specified use cases was found", "WARNING")
                            continue
                    else:
                        # Se use_cases estiver vazia, processa todos os use cases encontrados
                        pastas_para_processar = list(pastas_use_cases.keys())
                    
                    # Processa os use cases
                    for use_case in pastas_para_processar:
                        clicar_pasta(use_case, pastas_use_cases)
                        

                        pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
                        # Usa valores padrão se não encontrou a imagem
                        if pos_x is None:
                            pos_x, pos_y = 0.3, 0.1  # Valores padrão
                            registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
                        time.sleep(0.5)  
                        sub_pastas = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
                        vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)

                        # Se não encontrou nem subpastas nem VFs, é uma pasta vazia
                        if not sub_pastas and not vf_nomes:
                            registrar_caminho(
                                projeto=projeto,
                                pasta_nivel=pasta_maior_nivel,
                                pasta_requisitos=pasta_requisitos,
                                dominio=dominio_encontrado,
                                pasta_use_case=use_case,
                                pasta_vazia=True,
                                vfs_list=VFs
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
                                baixada=False if baixar else None,  # None para VFs que não precisam ser baixadas
                                vfs_list=VFs
                            )
                            
                            if baixar:
                                clicar_pasta(vf_nome, vf_nomes)
                                esperarPor(["abrir_somente_leitura.png", "abrir_somente_leitura_en.png"], timeout=30, iniX=0.05, iniY=0.05, fimX=0.8, fimY=0.95)
                                moveAndClick(["abrir_somente_leitura.png", "abrir_somente_leitura_en.png"], "left")
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
                                    baixada=sucesso,
                                    vfs_list=VFs
                                )

                        if sub_pastas != {}:
                            for sub_pasta in sub_pastas:
                                clicar_pasta(sub_pasta, sub_pastas)
                                pos_x, pos_y = encontrar_posicao_xy(["tipo_menu.png", "tipo_menu_en.png"], iniX=0.1, iniY=0.05, fimX=0.8, fimY=0.4)
                                # Usa valores padrão se não encontrou a imagem
                                if pos_x is None:
                                    pos_x, pos_y = 0.3, 0.1  # Valores padrão
                                    registrar_log("Could not find tipo_menu.png, using default coordinates", "WARNING")
                                    
                                vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=pos_y, fimX=pos_x, fimY=0.95)
                                
                                # Se não encontrou VFs na subpasta, é uma pasta vazia
                                if not vf_nomes:
                                    registrar_caminho(
                                        projeto=projeto,
                                        pasta_nivel=pasta_maior_nivel,
                                        pasta_requisitos=pasta_requisitos,
                                        dominio=dominio_encontrado,
                                        pasta_use_case=use_case,
                                        sub_pasta=sub_pasta,
                                        pasta_vazia=True,
                                        vfs_list=VFs
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
                                        baixada=False if baixar else None,  # None para VFs que não precisam ser baixadas
                                        vfs_list=VFs
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
                                            baixada=sucesso,
                                            vfs_list=VFs
                                        )

                                time.sleep(1)
                                voltar_nivel(1)

                            time.sleep(1)   
                            voltar_nivel(2)
                        else:
                            time.sleep(1)
                            voltar_nivel(1)
            else:
                # Lista os domínios encontrados no mapa
                dominios_encontrados = [nome_pasta for nome_pasta in pastas_dominios.keys()]
                registrar_log(f"None of the specified domains was found in project {projeto}. Available domains: {', '.join(dominios_encontrados)}", "WARNING")
                voltar = 4
        else:
            registrar_log("Failed to map domain folders", "ERROR")
            messagebox.showerror("Timeout", "Image recognition failed")
            return
        
        voltar_nivel(voltar)
        moveAndClick("projects.png", "left")
        time.sleep(0.5)

    print(f"\n✅ Process completed! The spreadsheet was saved at: {nome_arquivo_vfs}")
    messagebox.showinfo("Completed", f"Process finished!\nThe spreadsheet was saved at:\n{nome_arquivo_vfs}")
    return nome_arquivo_vfs

# Código removido de inicialização direta:
# projetos = ["332BEV"]
# dominios = ["Climate", "Comfort Climate"]
# use_cases = ["Defroster"]
# VFs = ["VF126"]

# O código agora é executado através da função main_logic chamada pela GUI