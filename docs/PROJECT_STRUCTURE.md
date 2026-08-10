# Estrutura do Projeto

Este documento registra a organizacao esperada do repositorio e ajuda a evitar
que notebooks, dados locais e scripts de dominio se misturem na raiz.

## Raiz

- `README.md`: entrada principal para setup e operacao.
- `requirements.txt`: dependencias principais em Windows.
- `requirements_mac.txt`: dependencias alternativas para macOS.
- `config.py`: leitura de variaveis de ambiente para Azure Batch/Storage.
- `.env.example`: modelo de configuracao local.

## Codigo

- `src/anexo_a`: processamento do Anexo A e criacao de base reprocessada.
- `src/anexo_b`: processamento do Anexo B.
- `src/anexo_c`: validacao de planilhas, dashboard e ajuste de contas.
- `src/resumo`: quadros de IRPJ/PIS e processamento consolidado de resumo.
- `src/export_kits`: empacotamento e organizacao de PDFs para envio.
- `src/azure`: integracoes com Azure Batch e Storage.
- `src/utils`: funcoes e classes compartilhadas.

## Documentos e Insumos Versionados

Use `docs` para arquivos pequenos e estaveis que fazem parte do projeto:

- templates `.xlsx`;
- listas auxiliares `.csv`;
- documentacao tecnica;
- exemplos sanitizados.

Nao versionar bases reais, entregas, dumps, PDFs gerados ou arquivos temporarios
do Excel.

## Notebooks

Notebooks sao uteis para execucao assistida e investigacao, mas devem chamar
funcoes de `src` em vez de concentrar regra de negocio.

Destino recomendado:

- `notebooks/processamento`: notebooks operacionais.
- `notebooks/exploracao`: testes manuais e investigacoes.

Os notebooks ainda existentes na raiz podem ser migrados quando nao houver
execucao em andamento ou referencias externas aos caminhos antigos.

## Testes

`tests` deve cobrir regras de negocio reutilizaveis de `src`.

Comando padrao:

```powershell
python -m pytest
```
