**SISTEMA AUTOMATIZADO DE ALERTA DE VENCIMENTO DE PROCURAÇÕES**

**Objetivo:**

Automatizar o controle de vencimento de procurações jurídicas, reduzindo riscos operacionais e eliminando a dependência de controle manual, por meio de alertas automatizados baseados em criticidade e prazo.

---
**Contexto**

O controle de vencimento de procurações era realizado manualmente, o que aumentava o risco de prazos não monitorados e dependência de acompanhamento individual.

**Regras de Negócio**

- 🔴 ALTA criticidade → alerta até 45 dias
- 🟡 MÉDIA e 🟢 BAIXA → alerta até 30 dias

**Classificação visual**

- 🚨 URGENTE → até 7 dias
- ⚠️ ATENÇÃO → até 15 dias  
- ❌ VENCIDO → prazo expirado

---

**Tecnologias Utilizadas**

- Python
- Pandas
- HTML/CSS (para formatação do e-mail)
- Outlook (automação de envio)
- Windows Task Scheduler (execução automática)

---

**Funcionalidades**

- Leitura de dados via Excel
- Cálculo automático de dias restantes
- Aplicação de regras de criticidade
- Geração de e-mail HTML com destaque visual
- Envio automático de alertas
- Execução diária automatizada

---

**Fluxo do Sistema**

1. Leitura da base em Excel  
2. Tratamento e padronização dos dados  
3. Cálculo de dias restantes  
4. Aplicação das regras de negócio  
5. Geração do e-mail em HTML  
6. Envio automático via Outlook  
7. Execução diária via agendador  

---

**Impacto**

- Redução de risco de vencimentos não monitorados
- Eliminação de controle manual
- Apoio direto ao departamento jurídico

---

**Observação**

Este projeto foi desenvolvido com base em uma necessidade real do departamento jurídico, com foco em confiabilidade, automação e apoio à tomada de decisão.

---

**Exemplo de Saída**

<img width="991" height="477" alt="image" src="https://github.com/user-attachments/assets/c8e1b43b-d586-45f7-aa6e-c2a5a0e388bb" />

**Diferenciais**

- Aplicação de regra de negócio real
- Automação de processo manual
- Integração com ambiente corporativo
- Comunicação visual otimizada para tomada de decisão

**Próximos Passos**

- Desenvolvimento de dashboard em Power BI
- Controle de histórico de alertas
- Evolução para integração com banco de dados

