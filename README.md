# Pricing BRAP — Automação de Precificação e Abertura de Chamados

> ⚠️ **Projeto em desenvolvimento.** Etapas 1 e 2 (coleta/enriquecimento e
> precificação interativa) estão funcionais. Exportação e abertura de
> chamado (RPA) ainda não implementadas.

## O que é

Automatiza o fluxo de precificação de exames sob demanda que hoje é feito
manualmente: coleta as pendências do controle mestre, enriquece com dados
de custo e tabela padrão, permite a decisão de preço de forma interativa
e rastreável, e (em breve) gera o arquivo de repasse para representantes
e abre os chamados automaticamente no portal, capturando o número de volta
para o controle.

O projeto nasceu de um processo 100% manual e repetitivo, e está sendo
reconstruído de forma incremental, com foco em portabilidade entre
máquinas do time e em não perder trabalho no meio de uma execução.

## Status das etapas

- [x] Coleta e enriquecimento de pendências (`dados_pricing.py`)
- [x] Precificação interativa com gravação imediata (`precificacao.py`)
- [ ] Exportação (arquivo por representante + arquivo consolidado)
- [ ] Abertura de chamado via RPA + captura automática do número

## Arquitetura

O controle mestre funciona como fonte única de verdade, com uma coluna de
`STATUS` que avança conforme a demanda é processada:

```
PENDENTE → PRECIFICADO → CHAMADO_ABERTO → CONCLUÍDO
```

Isso torna o processo retomável: se uma execução parar no meio, o que já
foi decidido continua salvo, e a próxima execução processa apenas o que
ainda está pendente — sem duplicar nem perder nada.

## Como rodar

1. Clone o repositório
2. Crie e ative o ambiente virtual:
   ```
   python -m venv .venv
   .venv\Scripts\activate.bat        # Windows (cmd)
   ```
3. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```
4. Instale o pacote em modo editável (necessário para os imports funcionarem):
   ```
   pip install -e .
   ```
5. Abra `notebooks/main.ipynb` e rode as células na ordem

## Pré-requisitos

- A biblioteca **"Pricing - Geral"** (Teams/SharePoint) precisa estar
  sincronizada localmente antes de rodar:
  `Teams → canal → Arquivos → Abrir no SharePoint → Sincronizar`
  (não use "Adicionar atalho" em arquivos individuais — isso cria um
  atalho de internet, não um arquivo sincronizado de verdade)
- Ajuste os caminhos em `src/pricing_brap/dados_pricing.py` (seção
  `CONFIG`) caso a estrutura de pastas sincronizada localmente seja
  diferente

## Estrutura do projeto

```
├── README.md
├── requirements.txt
├── pyproject.toml
├── .gitignore
├── src/
│   └── pricing_brap/
│       ├── __init__.py
│       ├── dados_pricing.py      # coleta e enriquecimento
│       └── precificacao.py       # loop interativo de precificação
├── notebooks/
│   └── main.ipynb                # ponto de entrada
└── tests/                        # a fazer
```

## Roadmap

- [ ] Módulo de exportação (arquivo tratado por representante + consolidado
      para a automação)
- [ ] Módulo de automação RPA (abertura de chamado + captura do número)
- [ ] Testes automatizados para as funções de localização de arquivo e
      validação de dados
- [ ] Regra de preço automática (quando a política de pricing for definida
      formalmente — hoje toda decisão é humana)

## Autor

Roger Gabriel D S F D Silva

## Licença

MIT — ver [LICENSE](LICENSE)
