
# Website Text Exporter

**Extraia textos visíveis de qualquer página e exporte para Excel — o texto na página é apagado para confirmar se tudo foi extraído.**

---

## O que é

Um utilitário Python que:

1. Conecta-se a uma instância do Microsoft Edge (depuração remota).  
2. Extrai **todos** os textos visíveis da página.  
3. Salva tudo num arquivo Excel (`.xlsx`).  
4. Remove os textos exportados da página (serve para o usuário confirmar oque foi extraído).  

---

## Pré-requisitos

- Windows (Edge instalado como navegador padrão)  
- Python 3.8+  
- Pacotes Python:
  ```bash
  pip install pandas selenium openpyxl
  ```
- Microsoft Edge iniciado com depuração remota:
  ```powershell
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
    --remote-debugging-port=9222 `
    --user-data-dir="C:\EdgeUserData"
  ```

---

## Instalação

1. Clone este repositório:  
   ```bash
   git clone https://github.com/seu-usuario/website-text-exporter.git
   cd website-text-exporter
   ```
2. Instale dependências:  
   ```bash
   pip install -r requirements.txt
   ```

---

## Uso

1. Inicie o Edge com o comando de depuração remota acima.  
2. Execute o script:
   ```bash
   python exporter.py
   ```
3. **Resultado**:
   - `website_texts.xlsx` com todos os textos capturados.  
   - Página web sem os textos que foram exportados.  

---

## Como funciona

- **Conexão**: usa Selenium + Edge DevTools Protocol na porta 9222.  
- **Extração**: o XPath `//*[normalize-space(text())]` captura todos os elementos com texto visível.  
- **Exportação**: `pandas` gera o `.xlsx`.  
- **Remoção**: JavaScript zera `textContent` dos elementos exportados.

---

## Dica rápida

> “Edge é o padrão em quase todo PC Windows — por isso aproveitei a depuração remota dele.”  

Esse detalhe está comentado no topo do código:  
```python
# "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeUserData"
```
