**SISTEMA DE ALERTA DE PROCURAÇÕES**

**Objetivo:**

Automatizar o controle de vencimento de procurações, enviando alertas por e-mail com base em criticidade e prazo, formulado e estruturado com base na necessidade real do departamento jurídico. 

---

**Regras de Negócio**

- Procurações de **ALTA criticidade**: alerta até 45 dias antes do vencimento
- Procurações de **MÉDIA e BAIXA criticidade**: alerta até 30 dias antes do vencimento 
- Destaque visual para:
  - URGENTE (até 7 dias)
  - ATENÇÃO (até 15 dias)
  - VENCIDO (inclui quantidade de dias desde que venceu)

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

**Impacto**

- Redução de risco de vencimentos não monitorados
- Eliminação de controle manual
- Apoio direto ao departamento jurídico

---

Exemplo de Saída

<img width="991" height="477" alt="image" src="https://github.com/user-attachments/assets/c8e1b43b-d586-45f7-aa6e-c2a5a0e388bb" />

**Diferenciais**

- Aplicação de regra de negócio real
- Automação de processo manual
- Integração com ambiente corporativo
- Comunicação visual otimizada para tomada de decisão

