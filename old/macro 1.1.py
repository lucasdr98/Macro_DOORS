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

def moveAndClick(image, clickType):
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
        center_x = x + w // 2
        center_y = y + h // 2
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
    #Mover o mouse para o canto superior esquerdo
    pyautogui.moveTo(10, 10)
    time.sleep(1)
    # Verifica se o arquivo de referência do ícone existe
    if not os.path.exists(icone_path):
        print(f"❌ Arquivo de ícone '{icone_path}' não encontrado!")
        return {}
    
    # Carrega a imagem do ícone e converte para escala de cinza
    icone_pasta = cv2.imread(icone_path, cv2.IMREAD_GRAYSCALE)
    if icone_pasta is None:
        print(f"❌ Erro ao carregar o ícone de pasta!")
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
    
    Args:
        nome_pasta: Nome da pasta a ser clicada
        mapa_pastas: Dicionário com mapeamento de pastas
        
    Returns:
        bool: True se encontrou e clicou, False caso contrário
    """

        
    print(f"🖱️ Buscando pasta '{nome_pasta}'")
    
    if mapa_pastas is None or len(mapa_pastas) == 0:
        print("❌ Não há mapeamento de pastas disponível!")
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
        print(f"✅ Pasta '{pasta_texto}' encontrada no mapeamento")
        
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
        print(f"❌ Pasta '{nome_pasta}' não encontrada no mapeamento.")
        print("Dica: Verifique se a pasta está visível na árvore ou atualize o mapeamento.")
        return False

def esperarPor(image, timeout=10, iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95):
    """
    Espera pela aparição de uma imagem em uma região específica da tela.
    
    Args:
        image: Nome do arquivo de imagem a ser procurado na pasta 'images/'
        timeout: Tempo máximo de espera em segundos
        iniX, iniY, fimX, fimY: Coordenadas relativas da região de busca
        
    Returns:
        bool: True se a imagem foi encontrada, False caso contrário
    """
    start_time = time.time()
    
    # Carrega a imagem de referência
    template = cv2.imread(r"images/"+image, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print(f"❌ Erro ao carregar imagem '{image}'. Verifique se existe em 'images/'")
        return False
    
    while time.time() - start_time < timeout:
        # Captura uma captura de tela
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
        regiao_gray = cv2.cvtColor(regiao, cv2.COLOR_RGB2GRAY)
        
        if debug:
            cv2.imwrite(f"{debug_dir}/regiao_{image}.png", regiao_gray)
        # Usa correspondência de modelo para encontrar a imagem
        result = cv2.matchTemplate(regiao_gray, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        # Define um limite de similaridade
        threshold = 0.7
        if max_val >= threshold:
            time.sleep(1)
            return True
            
        time.sleep(1)
        
   
        
    print(f" Timeout de {timeout} segundos: '{image}' não encontrado.")
    return False

def baixarVF(nome_VF):
    #Organizar a VF
    moveAndClick("main.png", "right")
    time.sleep(1)
    moveAndClick("novo.png", "left")
    time.sleep(1)
    moveAndClick("barra.png", "left")
    time.sleep(1)
    moveAndClick("object_heading.png", "left")
    time.sleep(1)
    moveAndClick("inserir.png", "left")
    time.sleep(1)
    moveAndClick("barra.png", "left")
    time.sleep(1)
    moveAndClick("object_number.png", "left")
    time.sleep(1)
    moveAndClick("inserir.png", "left")
    time.sleep(1)
    moveAndClick("barra.png", "left")
    time.sleep(1)
    moveAndClick("object_level.png", "left")
    time.sleep(1)
    moveAndClick("inserir.png", "left")
    time.sleep(1)
    moveAndClick("barra.png", "left")
    time.sleep(1)
    moveAndClick("object_identifier.png", "left")
    time.sleep(1)
    moveAndClick("name.png", "left")
    time.sleep(1)
    pyautogui.write("RegID")
    time.sleep(1)
    moveAndClick("inserir.png", "left")
    time.sleep(1)
    moveAndClick("fechar.png", "left")
    time.sleep(1)
    moveAndClick("main.png", "right")
    time.sleep(1)
    moveAndClick("propriedades.png", "left")
    time.sleep(1)
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
    time.sleep(1)
    #Salvar o Excel
    timeout = esperarPor("excel_icone.png", timeout=1200, iniX= 0.1, iniY= 0.5, fimX= 0.90, fimY= 0.98)
    if not timeout:
        messagebox.showerror("Timeout", "O Excel não foi encontrado")
        exit()
    else:
       time.sleep(2)
       moveAndClick("excel_icone.png", "right")
       time.sleep(1)
       moveAndClick("close_excel.png", "left")
       time.sleep(3)
       moveAndClick("save_excel.png", "left")
       time.sleep(2)
       pyautogui.write(nome_VF)
       time.sleep(1)
       moveAndClick("save_excel2.png", "left")
       time.sleep(4)
       moveAndClick("close_VF.png", "left")
       time.sleep(2)
       moveAndClick("continuar_close_vf.png", "left")
       time.sleep(2)
       return True

def get_pasta_nivel(nome_pasta):
    """
    Determina o nível hierárquico de uma pasta.
    Work in Progress é o nível mais alto.
    Para as demais, o número determina o nível primário e a letra o nível secundário.
    """
    # Normaliza o nome da pasta para comparação
    nome_normalizado = nome_pasta.lower().strip()
    
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
    Encontra a pasta de maior nível hierárquico no mapa de pastas.
    """
    maior_nivel = (-1, -1)
    pasta_escolhida = None
    
    print("\nAnalisando níveis das pastas:")
    for nome_pasta in mapa_pastas.keys():
        nivel_atual = get_pasta_nivel(nome_pasta)
        print(f"Pasta: {nome_pasta} -> Nível: {nivel_atual}")
        if nivel_atual > maior_nivel:
            maior_nivel = nivel_atual
            pasta_escolhida = nome_pasta
    
    print(f"\nPasta escolhida: {pasta_escolhida} (Nível: {maior_nivel})")
    return pasta_escolhida

def encontrar_pasta_requisitos(mapa_pastas):
    """
    Encontra a pasta de requisitos funcionais entre as pastas mapeadas.
    Aceita várias variações do nome.
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
        nome_normalizado = nome_pasta.lower().strip()
        if any(termo in nome_normalizado for termo in termos_requisitos):
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
    pyautogui.press("tab")
    time.sleep(1)
    for i in range(nivel):
        pyautogui.hotkey('shift', 'left')  # Usa hotkey para pressionar shift + seta esquerda
        time.sleep(1)

projetos = ["846"]#["139EL","250MY24","250MY26","312MCA","332BEV","334MCA","356MCA","356MHEV","520MY24","637BEV","637MCA","846","965","ALFAMCA","ARM20","ARM23","LP3","M240","M240MY26-BEV","MASAHMCA","MASAHMY26","332TR"]
dominios = ["Comfort Climate"]
VFs = ["VF999"]

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
    
    # Mapeia as subpastas do projeto
    if esperarPor("pasta.png"):
        pastas_niveis = mapear_pastas(icone_path="images/pasta.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        
        # Encontra e clica na pasta de maior nível
        pasta_maior_nivel = encontrar_pasta_maior_nivel(pastas_niveis)
        if pasta_maior_nivel:
            print(f"Selecionando pasta de maior nível: {pasta_maior_nivel}")
            clicar_pasta(pasta_maior_nivel, pastas_niveis)
        else:
            print("❌ Nenhuma pasta válida encontrada")
            messagebox.showerror("Erro", "Nenhuma pasta válida encontrada no projeto")
            exit()
    else:
        print("❌ Parando")
        messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
        exit()

    # Procura e clica em Functional Requirements
    if esperarPor("pasta_amarela.png"):
        pastas_requerimentos = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        
        pasta_requisitos = encontrar_pasta_requisitos(pastas_requerimentos)
        if pasta_requisitos:
            if not clicar_pasta(pasta_requisitos, pastas_requerimentos):
                print("❌ Erro ao clicar na pasta de requisitos")
                messagebox.showerror("Erro", "Erro ao clicar na pasta de requisitos funcionais")
                exit()
        else:
            print("❌ Pasta de requisitos funcionais não encontrada")
            messagebox.showerror("Erro", "Pasta de requisitos funcionais não encontrada")
            exit()
    else:
        print("❌ Parando")
        messagebox.showerror("Timeout", "o reconhecimento de imagem falhou")
        exit()

    if esperarPor("pasta_amarela.png"):
        pastas_dominios = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
        for dominio in dominios:
            nome = dominio.replace(" ", "")   #Melhorar esse padrão de nome
            if nome in pastas_dominios:
                clicar_pasta(nome, pastas_dominios)

                if esperarPor("pasta_amarela.png"):
                    pastas_use_cases = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                    print(pastas_use_cases)
                    for use_case in pastas_use_cases:
                        clicar_pasta(use_case, pastas_use_cases)
                        time.sleep(1)

                        sub_pastas = mapear_pastas(icone_path="images/pasta_amarela.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)

                        vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                        # Adiciona todas as VFs encontradas à planilha
                        for vf_nome, vf_info in vf_nomes.items():
                            df_vfs = adicionar_vf_planilha(
                                df_vfs,
                                nome_arquivo_vfs,
                                folder=use_case,
                                vf_info=vf_info,
                                projeto=projeto
                            )
                            
                            # Continua com o processo de download se necessário
                            for vf_esperado in VFs:
                                if vf_nome.lower().startswith(vf_esperado.lower()):
                                    clicar_pasta(vf_nome, vf_nomes)
                                    time.sleep(1)
                                    moveAndClick("abrir_somente_leitura.png", "left")
                                    esperarPor("main.png", timeout=20, iniX=0.1, iniY=0.1, fimX=0.9, fimY=0.5)
                                    baixarVF(vf_nome)

                        if sub_pastas != {}:
                            for sub_pasta in sub_pastas:
                                clicar_pasta(sub_pasta, sub_pastas)
                                vf_nomes = mapear_pastas(icone_path="images/icone_vf.png", iniX=0.1, iniY=0.1, fimX=0.23, fimY=0.95)
                                # Adiciona todas as VFs encontradas à planilha
                                for vf_nome, vf_info in vf_nomes.items():
                                    df_vfs = adicionar_vf_planilha(
                                        df_vfs,
                                        nome_arquivo_vfs,
                                        folder=f"{use_case}/{sub_pasta}",
                                        vf_info=vf_info,
                                        projeto=projeto
                                    )
                                    
                                    # Continua com o processo de download se necessário
                                    for vf_esperado in VFs:
                                        if vf_nome.lower().startswith(vf_esperado.lower()):
                                            clicar_pasta(vf_nome, vf_nomes)
                                            time.sleep(1)
                                            moveAndClick("abrir_somente_leitura.png", "left")
                                            esperarPor("main.png", timeout=20, iniX=0.1, iniY=0.1, fimX=0.9, fimY=0.5)
                                            baixarVF(vf_nome)
                                time.sleep(1)
                                voltar_nivel(1)

                            time.sleep(1)   
                            voltar_nivel(2)
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

