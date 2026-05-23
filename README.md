# AuditFiscal Pro

Sistema de Auditoria e Conferência Fiscal / Estoque

## Instalação

```bash
pip install -r requirements.txt
```

## Execução

```bash
python main.py
```

## Estrutura dos Arquivos Excel

### 1. Arquivo de Entrada de Estoque
Colunas necessárias (nomes aproximados, o sistema detecta automaticamente):
- **Documento** — número da nota fiscal
- **Fornecedor** — nome do fornecedor
- **Data Emissão** — data de emissão da nota
- **Total R$** — valor total da nota

### 2. Arquivo de Notas Emitidas (ForLions)
- Cabeçalho **obrigatoriamente na linha 5** do Excel
- Colunas utilizadas: Número, CNPJ, Razão Social, Dt. Emissão, Valor (Total/Valor), Cancelada

## Funcionalidades

- **Dashboard**: KPIs em tempo real com percentual de conformidade
- **Importar Arquivos**: Upload dos dois Excel, log em tempo real, barra de progresso
- **Resultados**: Tabela filtável com todas as notas (conciliadas, divergências, despesas, devoluções)
- **Divergências**: Classificação manual de cada nota sem entrada (Erro, Despesa, Devolução, Serviço...)
- **Empresas**: Cadastro de CNPJs conhecidos para classificação automática
- **Histórico**: Registro de todas as auditorias anteriores
- **Exportar Excel**: Relatório profissional com abas separadas e formatação automática

## Banco de Dados

O sistema usa SQLite local (`audit.db`) criado automaticamente na primeira execução.

## Arquitetura

```
fiscal_audit/
├── main.py              # Interface principal (CustomTkinter)
├── database/db.py       # Acesso ao banco SQLite
├── services/comparator.py  # Motor de comparação
├── exports/excel_export.py # Geração do relatório Excel
└── requirements.txt
```
