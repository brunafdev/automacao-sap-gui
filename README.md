# 🖱️ Bot de Automação SAP GUI

> **Status:** Concluído (Legado)
> **Contexto:** Solução desenvolvida para automatizar processos em ambiente onde o **SAP Scripting estava bloqueado**.

### O Desafio
A equipe precisava consultar dados bancários de centenas de fornecedores diariamente. O processo era 100% manual e repetitivo.
A ferramenta oficial de automação (SAP Scripting API) ainda não estava habilitada para o perfil de usuário, impedindo o uso de scripts VBA ou Python diretos.

### 💡 A Solução: RPA Visual
Como alternativa, desenvolvi um robô baseado em **reconhecimento de imagem** que simula a interação humana (mouse e teclado).

**Fluxo da Automação:**
1.  **Leitura:** O script lê uma lista de códigos de fornecedores em Excel (`openpyxl`).
2.  **Navegação:** Interage com o SAP GUI usando atalhos de teclado e cliques simulados.
3.  **Visão:** Utiliza a biblioteca `PyAutoGUI` para localizar campos na tela (ex: "Aba Dados Bancários") através de *printscreens* de referência.
4.  **Extração:** Copia os dados para a área de transferência e salva na planilha.

```mermaid
graph TD
    A[📄 Início: Ler Planilha Excel] -->|Loop por Linha| B{Tem código?}
    B -- Não --> Z[💾 Salvar e Finalizar]
    B -- Sim --> C[📋 Copiar Código ERP]
    C --> D[👁️ PyAutoGUI: Localizar Campo no SAP]
    D --> E[🖱️ Simular Cliques e Atalhos]
    E --> F[📥 Extrair Dados Bancários]
    F --> G[📝 Gravar no Excel]
    G --> B

### ⚠️ Limitações Conhecidas & Aprendizados
Por ser uma automação baseada em coordenadas visuais (pixels e imagens), esta solução possui dependências:
* **Resolução de Tela:** O robô depende da resolução do monitor ser mantida.
* **Interrupções:** Pop-ups inesperados do sistema podem pausar o fluxo.

> *Nota: Atualmente, recomendo o uso do **SAP Scripting nativo** (quando disponível) para maior robustez e precisão (100%), eliminando a margem de erro visual.*

### Tecnologias
* **Python 3.x**
* **PyAutoGUI** (Controle de Mouse/Teclado e Busca de Imagem)
* **OpenPyXL** (Manipulação de Excel)
* **PyPerClip** (Gestão de Clipboard)
