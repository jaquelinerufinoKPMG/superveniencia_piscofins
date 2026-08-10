# Superveniencia CSLL / PIS Cofins

Projeto Python para processar anexos, quadros e resumos relacionados a
superveniencia, CSLL, IRPJ, PIS e Cofins.

## Visao Geral

O repositorio esta organizado por dominio de processamento:

- `src/anexo_a`: rotinas do Anexo A e reprocessamento da base.
- `src/anexo_b`: rotinas do Anexo B.
- `src/anexo_c`: validacoes, dashboards e atualizacao de nomes de contas.
- `src/resumo`: geracao de quadros e processamento de resumos.
- `src/export_kits`: organizacao de PDFs e montagem de kits de envio.
- `src/azure`: rotinas de Azure Batch/Storage.
- `src/utils`: utilitarios compartilhados.
- `tests`: testes automatizados.
- `docs`: templates, listas auxiliares e documentacao de apoio.
- `notebooks`: destino recomendado para notebooks exploratorios e de processamento.

## Setup

Crie e ative um ambiente virtual:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Instale as dependencias:

```powershell
pip install -r requirements.txt
```

Em macOS, use o arquivo alternativo quando necessario:

```bash
pip install -r requirements_mac.txt
```

## Configuracao

As credenciais de Azure Batch/Storage sao lidas por variaveis de ambiente em
`config.py`.

Copie `.env.example` para `.env` e preencha os valores locais:

```powershell
Copy-Item .env.example .env
```

Variaveis esperadas:

- `BATCH_ACCOUNT_NAME`
- `BATCH_ACCOUNT_KEY`
- `BATCH_ACCOUNT_URL`
- `STORAGE_ACCOUNT_NAME`
- `STORAGE_ACCOUNT_KEY`
- `STORAGE_ACCOUNT_DOMAIN`

## Testes

Execute a suite com:

```powershell
python -m pytest
```

## Convencoes

- Codigo reutilizavel deve ficar em `src`, separado por dominio.
- Testes devem ficar em `tests`, acompanhando o comportamento dos modulos.
- Templates, listas e documentos de referencia ficam em `docs`.
- Notebooks devem ser usados como camada de execucao/analise, chamando funcoes de
  `src` sempre que possivel.
- Dados de entrada, saida e entregas locais nao devem ser versionados; use `data`,
  `Input`, `Output` ou `PRONTO`.

Veja tambem [docs/PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md).

