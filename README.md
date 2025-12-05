# 🚀 High-Performance Sales Suite for Google Sheets

> **Um ecossistema CRM modular desenvolvido para operar com latência zero dentro do Google Sheets.**

Muitas soluções de CRM integradas ao Google Sheets falham por serem "pesadas", causando travamentos e lentidão no navegador. Este projeto resolve esse problema através de uma **Arquitetura Desacoplada**.

Ao invés de uma aplicação monolítica, o sistema é dividido em dois módulos independentes e leves, garantindo que o vendedor tenha velocidade máxima seja na gestão de dados ou na comunicação.

## 🏗️ Decisão de Arquitetura & Performance

O Google Apps Script renderiza interfaces via *Iframes*. Interfaces complexas tendem a sobrecarregar o thread principal da planilha.
Por isso, este projeto foi separado intencionalmente em dois contextos:

1.  **Módulo CRM (Gestão):** Focado em operações de banco de dados (CRUD), filtros e status.
2.  **Módulo Comunicador (Disparo):** Focado em APIs externas (WhatsApp/Gmail) e limpeza de strings.

**Resultado:** O usuário pode manter o CRM aberto para gestão sem que o carregamento de scripts de comunicação afete a fluidez da planilha, e vice-versa.

## 🛠️ Módulos do Sistema

### 1. 📝 Módulo Gestão (CRM Sidebar)
Painel lateral dedicado ao ciclo de vida do cliente.
* **Smart Forms:** Validação de entrada e categorização por nicho.
* **Gestão de Pipeline:** Atualização rápida de Status (Prospecção -> Ganho/Perdido) e Próximos Passos.
* **Organização Automática:** Scripts que reordenam a planilha e arquivam leads finalizados.

### 2. ⚡ Módulo Comunicador (Quick Connect)
Interface leve para disparo de mensagens, eliminando o "copia e cola".
* **Busca & Autopreenchimento:** Localiza o lead na base e preenche os campos de contato instantaneamente.
* **WhatsApp API Engine:** Higieniza números de telefone, corrige DDI (+55) automaticamente e abre a conversa.
* **Disparador de E-mail:** Envia mensagens transacionais usando a infraestrutura do Gmail.

## 💻 Tecnologias Utilizadas

* **Front-end:** HTML5, CSS3 (Material Design Leve), JavaScript Vanilla.
* **Back-end:** Google Apps Script (GAS) Server-side processing.
* **Integração:** `SpreadsheetApp`, `MailApp`, WhatsApp Web Intent.

---
*Este projeto demonstra como superar as limitações de performance do Google Sheets através de código limpo e segregação de responsabilidades.*
