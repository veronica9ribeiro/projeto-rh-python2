# projeto-rh-python2


# Projeto M3 — Datas, fusos e cálculos de RH

Projeto prático do módulo M3 do Roadmap Python RH — Saint-Gobain.  
Lê uma planilha de colaboradores, aplica cálculos de RH baseados em datas e fusos horários, e gera um relatório `.xlsx` formatado com duas abas: **Relatório** e **Alertas**.

---

## Funcionalidades

- Conversão de datas da planilha (string → datetime) com `pandas`
- Cálculo de tempo de casa (anos e meses)
- Cálculo de aviso prévio conforme CLT (30 dias + 3 por ano, máx. 90)
- Período aquisitivo de férias e detecção de férias vencidas
- Registro de ponto localizado por fuso horário da UF com `pytz`
- Cálculo de dias úteis no mês com `numpy`
- Exportação de relatório `.xlsx` formatado com `openpyxl`

---

## Tecnologias

| Biblioteca | Uso |
|---|---|
| `pandas` | Leitura, manipulação e exportação de dados |
| `openpyxl` | Formatação do relatório Excel |
| `pytz` | Localização e conversão de fusos horários |
| `numpy` | Cálculo de dias úteis (`busday_count`) |

---

## Estrutura

```
projeto-rh-python2/
├── Projeto_M3_Datas_fusos_e_cálculos_de_RH.py   # script principal
├── colaboradores.xlsx                             # planilha de entrada (gerada pelo script)
├── relatorio_rh.xlsx                              # relatório de saída
└── README.md
```

---

## Como usar

**1. Instalar dependências**
```bash
pip install pandas openpyxl pytz numpy
```

**2. Rodar o script**
```bash
python Projeto_M3_Datas_fusos_e_cálculos_de_RH.py
```

O script gera automaticamente a planilha de entrada com dados fictícios e exporta o relatório `relatorio_rh.xlsx`.

**3. Para usar com dados reais**  
Substitua a chamada `criar_planilha_entrada()` pela leitura do seu arquivo:
```python
df = ler_colaboradores("seu_arquivo.xlsx")
```
A planilha precisa ter as colunas: `matricula`, `nome`, `cargo`, `uf`, `tipo`, `admissao`, `salario`, `registro_ponto`.

---

## Saída

**Aba Relatório** — todos os colaboradores com:
- Tempo de casa, aviso prévio, período aquisitivo, status do ponto
- Linhas alternadas em azul claro; alertas destacados em amarelo

**Aba Alertas** — filtra apenas colaboradores com férias vencidas ou atraso no ponto

---

## Contexto

Módulo M3 do Roadmap Python RH — Saint-Gobain.  
Módulos anteriores: M1 (base mínima), M2 (pandas/openpyxl).  
Próximo: M4 (integração com APIs via `requests` e OAuth2).
