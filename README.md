# 📊 Planilha Diego (Web App em HTML)

> **Aplicativo 100% em HTML e JavaScript para Automação de Curva ABC & Dashboard de Inventário**

O **Planilha Diego** foi desenvolvido em **HTML5 puro (Client-Side)** para rodar diretamente no navegador, sem necessidade de servidores complexos. É perfeito para ser hospedado gratuitamente no **GitHub Pages** ou executado offline dando dois cliques no arquivo `index.html`.

---

## 🌟 O que o Aplicativo Faz?

1. **Anexo Simples:** Você anexa a planilha bruta simples e desorganizada baixada do sistema (ex: `INVENTARIO...XLSX`).
2. **Processamento Instantâneo:** O aplicativo lê e calcula tudo em segundos na memória do próprio navegador:
   - Unifica as frentes e almoxarifados na aba **Base**.
   - Agrupa os produtos e calcula a **Curva ABC** (A: até 80%, B: até 95%, C: restante) com formatação condicional colorida.
   - Monta as tabelas de **Análises** (classes, locais, naturezas, top itens, matriz).
   - Monta o **Dashboard** com 6 cards de KPI e gráficos.
3. **Dashboard Interativo na Tela:** Exibe na tela do navegador os indicadores em tempo real e 7 gráficos interativos (Chart.js):
   - Curva ABC Acumulada
   - Pareto (20 maiores itens)
   - Rosca de Valor por Classe ABC
   - Itens por Classe ABC
   - Valor por Local de Estoque
   - Top 10 Naturezas
   - Top 10 Itens por Valor
4. **Download Imediato (.xlsx):** Gera automaticamente a planilha Excel pronta para download com todas as 7 abas, fórmulas dinâmicas (`INDEX/MATCH`, `SUMIF`, `COUNTIF`), cores e totais.

---

## 🚀 Como Usar

### Opção 1: Abrir no Navegador (Offline com 2 Cliques)
Basta dar um duplo clique no arquivo:
```
index.html
```
Ele abrirá no seu navegador padrão (Google Chrome, Microsoft Edge, etc.) pronto para uso imediato!

---

### Opção 2: Hospedar no GitHub Pages (Online 24h Grátis)
Como o aplicativo é em **HTML**, você pode hospedá-lo no GitHub Pages gratuitamente para que qualquer pessoa da empresa possa usar pelo link:

1. Suba o projeto para o seu repositório Git:
   ```bash
   git init
   git add .
   git commit -m "feat: Aplicativo Planilha Diego em HTML"
   git branch -M main
   git remote add origin https://github.com/SEU_USUARIO/planilha-diego.git
   git push -u origin main
   ```
2. No seu repositório no GitHub:
   - Vá em **Settings** > **Pages** (no menu lateral esquerdo).
   - Em **Source**, selecione a branch `main` e a pasta `/(root)`.
   - Clique em **Save**.
3. Em 1 minuto o GitHub fornecerá o link público (ex: `https://seu-usuario.github.io/planilha-diego/`).

---

### Opção 3: Executar pelo Batch do Windows
Você também pode clicar duas vezes em:
```
INICIAR_APLICATIVO.bat
```
Ele iniciará a visualização local no navegador em `http://localhost:5000`.

---

## 📁 Estrutura de Arquivos

```
PLANILHA DI/
├── index.html               # Aplicativo Web completo em HTML/JS
├── INICIAR_APLICATIVO.bat   # Inicializador para Windows (opcional)
├── app.py                   # Servidor Flask opcional para execução local
├── excel_processor.py       # Motor auxiliar de processamento
├── requirements.txt         # Dependências auxiliares
├── Procfile                 # Deploy em nuvem (opcional)
├── .gitignore               # Configurações do Git
└── README.md                # Este manual explicativo
```

---

## 🛠️ Tecnologias Utilizadas

- **HTML5 & CSS3** (Interface moderna e responsiva)
- **Tailwind CSS** (Design elegante e profissional)
- **Lucide Icons** (Ícones modernos)
- **SheetJS (xlsx.full.min.js)** (Leitura rápida de arquivos Excel no navegador)
- **ExcelJS** (Geração da planilha Excel com estilos, fórmulas e cores)
- **Chart.js** (Gráficos interativos no navegador)
