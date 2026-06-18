# 📊 Radar de Implantação

Automação desenvolvida em **Excel VBA** para geração automática de apresentações executivas em **PowerPoint**, transformando dados de projetos em dashboards gerenciais prontos para reuniões e comitês executivos.

---

## 🎯 Objetivo

Reduzir o esforço manual na criação de relatórios executivos, consolidando informações de projetos e gerando apresentações padronizadas de forma automática.

---

## 🚀 Principais Funcionalidades

✅ Leitura automática da tabela `tblProjects`

✅ Cálculo consolidado de KPIs

✅ Geração automática de apresentações PowerPoint

✅ Paginação automática para grandes volumes de projetos

✅ Destaque visual para indicadores críticos

✅ Formatação executiva padronizada

✅ Criação automática de tabelas gerenciais

---

## 📈 Indicadores Gerados

A automação calcula e apresenta:

| Indicador | Descrição |
|------------|------------|
| Total de Projetos | Quantidade total de projetos monitorados |
| Total de Pendências | Soma das pendências abertas |
| Taxa Global de Pendências | Percentual consolidado de criticidade |
| Dias para Go-Live | Prazo restante para implantação |
| Status RH | Situação de contratação/alocação |
| Status Operacional | Situação operacional do projeto |
| Data de Go-Live | Data prevista de lançamento |

---

## 🛠️ Tecnologias Utilizadas

- Microsoft Excel 365
- VBA (Visual Basic for Applications)
- Microsoft PowerPoint Object Model
- ListObjects (Tabelas Estruturadas)
- Automação COM

---

## 🔄 Fluxo da Solução

```text
Base de Dados (Excel)
          ↓
Processamento VBA
          ↓
Cálculo dos KPIs
          ↓
Geração Automática do PowerPoint
          ↓
Dashboard Executivo
```

---

## 📋 Estrutura da Base

| Campo | Descrição |
|---------|------------|
| BUSINESS_UNIT | Unidade responsável |
| PROJECT_ID | Identificador único |
| PROJECT_NAME | Nome do projeto |
| CITY | Cidade |
| STATE | Estado |
| CREATED_AT | Data de criação |
| TEAM_SIZE | Quantidade de recursos |
| HR_STATUS | Status de RH |
| OPS_STATUS | Status operacional |
| OPEN_ISSUES | Pendências abertas |
| ISSUE_RATE | Percentual de pendências |
| DAYS_TO_LAUNCH | Dias restantes |
| LAUNCH_DATE | Data prevista de Go-Live |

---

## ⭐ Diferenciais

- Redução significativa do tempo de preparação de apresentações
- Padronização visual dos relatórios
- Eliminação de atividades manuais repetitivas
- Escalabilidade para dezenas ou centenas de projetos
- Fácil adaptação para PMO, Expansão, Operações e Implantações

---

## 💼 Casos de Uso

- Gestão de Portfólio de Projetos
- PMO (Project Management Office)
- Expansão de Operações
- Rollout de Unidades
- Acompanhamento de Implantações
- Relatórios Executivos

---

## 📷 Demonstração

### Dashboard

<img width="1894" height="783" alt="image" src="https://github.com/user-attachments/assets/7110611f-18b6-4fc7-84a6-fd2e5219b0d3" />


### Apresentação Gerada

<img width="1919" height="796" alt="image" src="https://github.com/user-attachments/assets/88c3642d-1b0a-47b3-894b-897cfa6261d3" />

---

## 👨‍💻 Autor

### Davi Lima

PMO | Planejamento | Automação de Processos

Desenvolvedor de soluções corporativas focado em automação de processos, análise de dados e produtividade utilizando Excel VBA, Power BI e PowerPoint Automation.

GitHub: https://github.com/dvblim

---

## 📄 Licença

Este projeto está licenciado sob a licença MIT.
