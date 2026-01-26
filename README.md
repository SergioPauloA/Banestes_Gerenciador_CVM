# 🚀 Monitor de Fundos CVM - BANESTES

<div align="center">

![Banestes Badge](https://img.shields.io/badge/Banestes-Investimentos-1e3a8a?style=for-the-badge)
![AppScript Badge](https://img.shields.io/badge/Google-AppScript-4285F4?style=for-the-badge)
![Status Badge](https://img.shields.io/badge/Status-Ativo-10b981?style=for-the-badge)
![Version Badge](https://img.shields.io/badge/Versão-4.0-3b82f6?style=for-the-badge)

**Uma solução completa e elegante para monitorar fundos de investimento em tempo real**

[📊 Dashboard](#-dashboard) • [📁 Estrutura](#-estrutura-do-projeto) • [⚙️ Funcionalidades](#-funcionalidades) • [🔧 Instalação](#-instalação) • [📚 Documentação](#-documentação)

</div>

---

## ✨ Visão Geral

O **Monitor de Fundos CVM - BANESTES** é um aplicativo web baseado em **Google AppScript** que oferece uma solução robusta e intuitiva para acompanhar o desempenho de **26 fundos de investimento BANESTES** em tempo real.

Desenvolvido com tecnologias modernas, este projeto integra coleta automática de dados, processamento de informações financeiras e apresentação visual avançada em um único painel de controle.

> ⚡ **Propósito:** Facilitar a tomada de decisão dos investidores através de dados precisos, atualizados e apresentados de forma clara e profissional.

---

## 🎯 Funcionalidades Principais

<details>
<summary><b>📊 1. Dashboard Interativo e Responsivo</b></summary>

- ✅ Interface web moderna e responsiva
- ✅ Gráficos e visualizações em tempo real
- ✅ Layout otimizado para desktop, tablet e mobile
- ✅ Tema visual profissional com gradientes e animações
- ✅ Painel de controle centralizado

</details>

<details>
<summary><b>💰 2. Monitoramento de 26 Fundos BANESTES</b></summary>

Acompanhe em tempo real:

- **Fundos de Renda Fixa Curto Prazo**
  - Banestes Investidor Automático
  - Banestes Invest Money
  - Banestes Solidez Automático

- **Fundos de Renda Fixa Referenciado DI**
  - VIP DI FIC
  - Vitória 500 FIC
  - Tesouro Referenciado
  - Valores
  - Liquidez
  - Reserva Climática

- **Fundos de Títulos Públicos**
  - Banestes IMA-B
  - Banestes IMA-B 5
  - Banestes IRF-M 1

- **Fundos de Ações**
  - Banestes BTG Pactual Absoluto Institucional
  - Banestes Dividendos
  - Banestes Tenax Ações
  - Banestes Synergy Long Only

- **Fundos Multimercado**
  - Banestes Funses
  - Banestes Multiestratégia

- **Fundos de Crédito Privado**
  - Banestes Selection
  - Banestes Crédito Corporativo I

- **Fundos Simples**
  - Banestes Invest Fácil
  - Banestes Soberano

- **Fundos Incentivados**
  - FIC Incentivados de Infraestrutura

- **Fundos Estratégicos**
  - Banestes Estratégia
  - Banestes Público Automático

</details>

<details>
<summary><b>🔄 3. Integração com CVM</b></summary>

- ✅ Leitura de dados oficiais da CVM (Comissão de Valores Mobiliários)
- ✅ Importação automática via `IMPORTXML`
- ✅ Dados sempre atualizados e confiáveis
- ✅ Códigos CVM para cada fundo

</details>

<details>
<summary><b>📅 4. Sistema de Datas e Calendário</b></summary>

- ✅ Reconhecimento automático de feriados brasileiros
- ✅ Cálculo de datas de referência inteligente
- ✅ Suporte para múltiplos anos (2025, 2026 e além)
- ✅ Integração com dias úteis bancários

</details>

<details>
<summary><b>📈 5. Análise de Performance</b></summary>

- ✅ Comparação de desempenho entre fundos
- ✅ Histórico de variações
- ✅ Indicadores de rentabilidade
- ✅ Análise comparativa com benchmarks

</details>

<details>
<summary><b>🛡️ 6. Segurança e Confiabilidade</b></summary>

- ✅ Conexão segura com Google Sheets
- ✅ Acesso restrito via autenticação
- ✅ Tratamento robusto de erros
- ✅ Logs detalhados de operações

</details>

---

## 📊 Dashboard

O dashboard oferece uma experiência completa com:

```
┌─────────────────────────────────────────────────────────────┐
│  🏦 Monitor de Fundos CVM - BANESTES v4.0                  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  📊 [Selecione um Fundo ▼]    🔄 Atualizar    📥 Exportar  │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Nome do Fundo: ___________________________          │  │
│  │  Código CVM: 275709                                 │  │
│  │  Data: 26/01/2026                                   │  │
│  │  Valor: R$ ___________                              │  │
│  │  Variação: +1,23%  ↗                                │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                             │
│  📈 [Gráfico de Performance]                               │
│  📊 [Comparativo de Fundos]                                │
│  📋 [Tabela Detalhada]                                     │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 📁 Estrutura do Projeto

```
Projeto - Banestes (4.0) - AppScript/
│
├── Code.gs                 # 🔧 Backend principal - Lógica de negócio
│   ├── Web App Handler (doGet)
│   ├── API de Dados (getDashboardData)
│   ├── Gerenciamento de Planilhas
│   └── Funções de Utilitários
│
├── ConfigData.gs           # ⚙️ Configurações e Dados Estáticos
│   ├── Feriados Brasileiros
│   ├── Constantes da Aplicação
│   └── Definições Globais
│
├── FundoService.gs         # 💰 Serviço de Fundos (26 fundos BANESTES)
│   ├── Lista de Fundos
│   ├── Códigos CVM
│   ├── Identifiers
│   └── Métodos de Acesso
│
├── DateUtils.gs            # 📅 Utilitários de Datas
│   ├── Cálculo de Datas de Referência
│   ├── Formatação de Datas
│   ├── Validação de Datas
│   └── Integração com Calendário
│
├── onInstall.gs            # 🚀 Script de Instalação
│   ├── Inicialização
│   ├── Configuração de Triggers
│   └── Setup Inicial
│
├── Index.html              # 🎨 Interface Frontend
│   ├── Layout Responsivo
│   ├── Estilos Modernos (CSS)
│   ├── JavaScript Interativo
│   ├── Dashboard Interativo
│   └── Efeitos Visuais
│
└── README.md              # 📖 Este arquivo
```

### Detalhes dos Arquivos

| Arquivo | Linhas | Descrição |
|---------|--------|-----------|
| **Code.gs** | ~1377 | Backend principal com APIs e lógica de negócio |
| **Index.html** | ~1298 | Frontend com interface interativa e responsiva |
| **FundoService.gs** | Múltiplo | 26 fundos BANESTES cadastrados e disponíveis |
| **ConfigData.gs** | ~107 | Feriados e configurações estáticas |
| **DateUtils.gs** | Múltiplo | Utilitários para cálculos de datas |
| **onInstall.gs** | - | Função de inicialização do AppScript |

---

## 🚀 Instalação

### Pré-requisitos

- ✅ Conta Google
- ✅ Acesso a Google Sheets
- ✅ Acesso a Google Apps Script
- ✅ Conexão com internet

### Passo a Passo

#### 1️⃣ **Clonar ou Copiar os Arquivos**

```bash
1. Copie todos os arquivos .gs para seu projeto AppScript
2. Copie o arquivo Index.html para a pasta HTML
3. Certifique-se de que o SPREADSHEET_ID está correto
```

#### 2️⃣ **Configurar o ID da Planilha**

No arquivo `Code.gs`, atualize:

```javascript
var SPREADSHEET_ID = '1N6LP1ydsxnQO_Woatv9zWEccb0fOGaV_3EKK1GoSWZI';
```

#### 3️⃣ **Publicar o Aplicativo**

```
Google Apps Script Editor → Deploy → New Deployment
Tipo: Web app
Executar como: Seu email
Quem tem acesso: Everyone
```

#### 4️⃣ **Acessar a Aplicação**

Copie o URL fornecido e acesse no navegador:

```
https://script.google.com/macros/d/{DEPLOYMENT_ID}/userweb
```

---

## ⚙️ Como Funciona

### 1. **Coleta de Dados**

O sistema utiliza `IMPORTXML` na planilha para coletar dados oficiais da CVM:

```javascript
function getDashboardData() {
  // Lê dados da planilha
  // Processa informações de 26 fundos
  // Retorna JSON estruturado
}
```

### 2. **Processamento**

Os dados são processados pelo backend AppScript:

```
Planilha (IMPORTXML) 
    ↓
Code.gs (Processamento)
    ↓
JSON (API)
    ↓
Index.html (Apresentação)
```

### 3. **Exibição**

O frontend apresenta os dados de forma visual:

- 📊 Gráficos interativos
- 📈 Tabelas responsivas
- 🎨 Tema moderno e elegante
- ⚡ Atualizações em tempo real

---

## 🎨 Recursos Visuais

### Tema e Cores

```
🎨 Paleta de Cores:
├── Primária: #1e3a8a (Azul Escuro)
├── Secundária: #3b82f6 (Azul Claro)
├── Sucesso: #10b981 (Verde)
├── Aviso: #f59e0b (Amarelo)
└── Erro: #ef4444 (Vermelho)
```

### Efeitos

- ✨ Gradientes suaves
- 🎭 Animações fluidas
- 📱 Responsividade completa
- 🖱️ Interatividade aprimorada
- 💫 Transições elegantes

---

## 📊 Dados dos Fundos

### Exemplo de Estrutura

```javascript
{
  nome: 'BANESTES INVESTIDOR AUTOMÁTICO FIF...',
  codigoCVM: '275709',
  codigoFundo: '1',
  data: '26/01/2026',
  valor: 10.523,
  variacao: 1.23,
  rentabilidade: 12.50
}
```

### Acesso aos Fundos

```javascript
// Obter todos os fundos
var fundos = getFundos();

// Total de fundos
var total = getTotalFundos(); // Retorna: 26
```

---

## 🔐 Segurança

✅ **Autenticação Google** - Requer login Google  
✅ **HTTPS** - Conexão criptografada  
✅ **Validação de Dados** - Verificação em tempo real  
✅ **Logs Auditáveis** - Rastreamento de operações  

---

## 📱 Compatibilidade

| Dispositivo | Status |
|------------|--------|
| 💻 Desktop | ✅ Totalmente suportado |
| 📱 Mobile | ✅ Responsivo e otimizado |
| 📲 Tablet | ✅ Interface adaptável |
| 🌐 Navegadores | ✅ Chrome, Firefox, Safari, Edge |

---

## 🚀 Casos de Uso

<details>
<summary><b>👨‍💼 Gerentes de Fundos</b></summary>

- Monitorar performance dos 26 fundos
- Gerar relatórios diários
- Acompanhar variações de preço

</details>

<details>
<summary><b>📊 Analistas Financeiros</b></summary>

- Comparar desempenho entre fundos
- Análise comparativa de rentabilidade
- Identificar tendências

</details>

<details>
<summary><b>💼 Investidores Institucionais</b></summary>

- Decisões de alocação baseadas em dados
- Acompanhamento em tempo real
- Histórico detalhado de performance

</details>

<details>
<summary><b>🎯 Equipe Comercial BANESTES</b></summary>

- Apresentações a clientes
- Comparativo de produtos
- Dados para propostas

</details>

---

## 🤝 Integração com Google Sheets

O projeto integra-se perfeitamente com Google Sheets:

```javascript
// Obter referência da planilha
var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// Acessar abas
var sheet = ss.getSheetByName('Fundos');

// Ler dados
var dados = sheet.getDataRange().getValues();
```

---

## 📈 Performance

- ⚡ Carregamento rápido
- 🔄 Atualizações eficientes
- 💾 Armazenamento otimizado
- 📊 Processamento escalável

---

## 🔧 Personalização

### Adicionar um Novo Fundo

No arquivo `FundoService.gs`:

```javascript
function getFundos() {
  return [
    // ... fundos existentes ...
    { 
      nome: 'NOVO FUNDO BANESTES', 
      codigoCVM: '999999', 
      codigoFundo: '99' 
    }
  ];
}
```

### Modificar Feriados

No arquivo `ConfigData.gs`:

```javascript
function getFeriadosBrasileiros() {
  return [
    [new Date(2026, 0, 1), 'quinta-feira', 'Confraternização Universal'],
    // ... adicione novos feriados ...
  ];
}
```

### Customizar Cores

No arquivo `Index.html`:

```css
:root {
  --primary-color: #1e3a8a;
  --secondary-color: #3b82f6;
  /* ... customize conforme necessário ... */
}
```

---

## 📚 Documentação de Funções

### Principais Funções

#### `doGet(e)`
- Retorna a interface web principal
- Ponto de entrada da aplicação

#### `getDashboardData()`
- Coleta dados de todos os 26 fundos
- Processa informações da planilha
- Retorna JSON estruturado

#### `obterPlanilha()`
- Abre a planilha configurada
- Trata erros de conexão

#### `getFundos()`
- Retorna lista de 26 fundos
- Inclui códigos CVM e identificadores

#### `getFeriadosBrasileiros()`
- Lista todos os feriados
- Suporta múltiplos anos

---

## 🐛 Solução de Problemas

### ❌ Erro: "Não foi possível abrir a planilha"

**Solução:**
- Verifique o `SPREADSHEET_ID`
- Confirme que tem acesso à planilha
- Verifique permissões no Google Drive

### ❌ Dados não aparecem

**Solução:**
- Atualize a página (F5)
- Verifique conexão com internet
- Confirme que as fórmulas IMPORTXML estão funcionando

### ❌ Interface com problemas de visualização

**Solução:**
- Limpe o cache do navegador
- Use navegador atualizado
- Teste em modo incógnito

---

## 📞 Suporte

Para dúvidas ou sugestões:

- 📧 Email: [seu-email]
- 💬 Slack: [canal-slack]
- 📱 WhatsApp: [seu-número]

---

## 📄 Licença

Este projeto é propriedade do BANESTES e seu uso está restrito conforme termos de licença.

```
© 2025 BANESTES
Todos os direitos reservados.
```

---

## 🎯 Roadmap Futuro

- [ ] 📱 App Mobile nativo (iOS/Android)
- [ ] 📊 Novos gráficos e análises
- [ ] 🤖 Previsões com IA
- [ ] 📧 Notificações por email
- [ ] 📱 Notificações push
- [ ] 🌐 Suporte a múltiplos idiomas
- [ ] 📈 Mais indicadores financeiros
- [ ] 🔔 Alertas personalizados
- [ ] 💾 Exportação avançada (PDF, Excel)
- [ ] 🗂️ Portfólio do usuário

---

## ✅ Checklist de Funcionalidades

- ✅ Monitor de 26 fundos em tempo real
- ✅ Dashboard responsivo e moderno
- ✅ Integração CVM
- ✅ Calendário de feriados
- ✅ APIs seguras
- ✅ Interface intuitiva
- ✅ Performance otimizada
- ✅ Tratamento de erros robusto
- ✅ Documentação completa
- ✅ Suporte a múltiplos dispositivos

---

<div align="center">

### 🌟 Desenvolvido com ❤️ para o BANESTES 🌟

![Made with Google Apps Script](https://img.shields.io/badge/Made%20with-Google%20Apps%20Script-4285F4?style=flat-square)
![JavaScript](https://img.shields.io/badge/Language-JavaScript-F7DF1E?style=flat-square)
![HTML5](https://img.shields.io/badge/HTML5-E34C26?style=flat-square)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=flat-square)

**[⬆ Voltar ao topo](#-monitor-de-fundos-cvm---banestes)**

</div>
