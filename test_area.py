import pyautogui
import cv2
import numpy as np
import time
import os

def testar_area_captura(iniX=0.4, iniY=0.4, fimX=1, fimY=0.80, duracao=5):
    """
    Testa a área de captura desenhando um retângulo na tela para visualização.
    
    Args:
        iniX, iniY: Coordenadas do ponto inicial (proporção da tela, 0 a 1)
        fimX, fimY: Coordenadas do ponto final (proporção da tela, 0 a 1)
        duracao: Tempo em segundos para manter a visualização
    """
    # Cria diretório de debug se não existir
    debug_dir = "debug_area"
    if not os.path.exists(debug_dir):
        os.makedirs(debug_dir)
    
    # Captura a tela
    screenshot = pyautogui.screenshot()
    screenshot_np = np.array(screenshot)
    screenshot_cv = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
    
    # Obtém dimensões da tela
    altura, largura = screenshot_cv.shape[:2]
    
    # Calcula coordenadas em pixels
    inicio_x = int(largura * iniX)
    fim_x = int(largura * fimX)
    inicio_y = int(altura * iniY)
    fim_y = int(altura * fimY)
    
    # Desenha retângulo vermelho na área de captura
    cv2.rectangle(screenshot_cv, 
                 (inicio_x, inicio_y), 
                 (fim_x, fim_y), 
                 (0, 0, 255), 2)
    
    # Adiciona texto com as coordenadas
    texto = f"Area: ({iniX:.2f}, {iniY:.2f}) -> ({fimX:.2f}, {fimY:.2f})"
    cv2.putText(screenshot_cv, texto,
                (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 
                1, (0, 0, 255), 2)
    
    # Salva a imagem original com o retângulo
    cv2.imwrite(f"{debug_dir}/area_completa.png", screenshot_cv)
    
    # Recorta e salva apenas a região selecionada
    regiao = screenshot_cv[inicio_y:fim_y, inicio_x:fim_x]
    cv2.imwrite(f"{debug_dir}/area_recortada.png", regiao)
    
    print(f"Dimensões da tela: {largura}x{altura}")
    print(f"Área selecionada em pixels: ({inicio_x}, {inicio_y}) -> ({fim_x}, {fim_y})")
    print(f"Tamanho da área: {fim_x - inicio_x}x{fim_y - inicio_y} pixels")
    print(f"\nImagens salvas em:")
    print(f"- {debug_dir}/area_completa.png (tela inteira com retângulo)")
    print(f"- {debug_dir}/area_recortada.png (apenas a área selecionada)")
    
    return inicio_x, inicio_y, fim_x, fim_y

if __name__ == "__main__":
    # Exemplo de uso
    print("Testando área de captura...")
    print("Pressione Ctrl+C para interromper")
    
    try:
        while True:
            # Você pode modificar estes valores para testar diferentes áreas
            testar_area_captura(iniX=0.1, iniY=0.05, fimX=0.7, fimY=0.5)
            time.sleep(2)  # Espera 2 segundos antes de atualizar
            
    except KeyboardInterrupt:
        print("\nTeste interrompido pelo usuário") 
        
        
        # icone doors (iniX=0.3, iniY=0.20, fimX=0.6, fimY=0.4)