# Macro DOORS

Automatização para interação com IBM Rational DOORS usando reconhecimento de imagem e OCR.

## Recursos

- Navegação automatizada pela interface do DOORS
- Reconhecimento de texto usando OCR (Tesseract)
- Extração e exportação de dados para Excel
- Suporte para múltiplas versões de interface (com detecção de múltiplas imagens)

## Requisitos

- Python 3.8+
- Tesseract OCR instalado no sistema (caminho padrão: `C:\Program Files\Tesseract-OCR\tesseract.exe`)
- Bibliotecas Python listadas em `requirements.txt`

## Instalação

1. Clone o repositório:
   ```
   git clone https://github.com/lucasdr98/Macro_DOORS.git
   ```

2. Crie e ative um ambiente virtual:
   ```
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

3. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

4. Instale o Tesseract OCR:
   - Baixe de: https://github.com/UB-Mannheim/tesseract/wiki
   - Instale no caminho padrão: `C:\Program Files\Tesseract-OCR\`

## Uso

Execute o script principal:
```
python macro.py
```

## Novos recursos

### Suporte para múltiplas versões de interface

As funções `moveAndClick` e `esperarPor` foram atualizadas para aceitar múltiplas imagens como alternativas:

```python
# Exemplo com uma única imagem (compatível com versões anteriores)
moveAndClick("botao.png", "left")

# Exemplo com múltiplas imagens alternativas
moveAndClick(["botao_v1.png", "botao_v2.png", "botao_v3.png"], "left")

# Da mesma forma para esperarPor
esperarPor(["dialogo_v1.png", "dialogo_v2.png"], timeout=10)
```

A função irá tentar encontrar qualquer uma das imagens fornecidas e usar a primeira que encontrar com maior confiança. 

## Compilação com PyInstaller

Para compilar o aplicativo em um executável standalone:

```
pyinstaller doors_macro.spec
```

Isso criará uma pasta `dist` com o executável e todos os arquivos necessários, incluindo o Tesseract OCR e as imagens usadas para reconhecimento. 