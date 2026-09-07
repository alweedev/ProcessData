# Arquitetura Modernizada - ProcessData

## Objetivo

Evoluir o projeto para uma arquitetura modular, legível e escalável, mantendo as regras de negócio atuais de cadastro, inativação e aprovação.

## Camadas

- `backend/domain`: regras e contratos de negócio.
- `backend/services`: orquestração de casos de uso.
- `backend/shared`: utilitários compartilhados e normalização.
- `backend/infra`: integrações técnicas e adaptação com arquivos/fontes.
- `backend/api`: rotas HTTP e serialização de entrada/saída.

## Mudanças aplicadas

- Criação de serviços dedicados para processamento, validação, exportação e relatório de qualidade.
- Endpoint de cadastro migrado para uso da camada de serviços.
- Novo endpoint de análise: `POST /api/analysis/summary`.
- Criação de testes para regras e normalizações críticas.
- Migração da lógica de inativação para `backend/services/inactivation_service.py`.
- Inclusão de trilha de auditoria persistente (JSONL) com API de histórico.
- Extração da feature de histórico para módulo próprio de frontend.

## Fluxos

### Cadastro
1. API recebe arquivos e valida extensão.
2. Service processa regras e normaliza dados.
3. Service exporta planilha final em Excel.
4. API retorna arquivo pronto para carga.

### Análise
1. API recebe arquivos e valida extensão.
2. Service processa os dados com regras atuais.
3. Service gera resumo de qualidade:
   - linhas totais
   - válidas
   - inválidas
   - duplicadas
   - vazios em campos críticos
4. API retorna relatório e preview para UI.

## Diretrizes de evolução

- Não concentrar novas regras em `processor.py`.
- Cada novo fluxo deve nascer em `domain + services + api`.
- Evitar arquivos únicos grandes no frontend; separar por feature.
- Manter cobertura mínima de testes para cada regra nova.

## Próxima fase

- Migrar gradualmente inativação e aprovação para serviços específicos.
- Continuar quebra do frontend em módulos por aba (`analise`, `cadastro`, `inativacao`, `aprovacao`, `historico`).
- Integrar a aba de análise com endpoint de resumo de qualidade para visão operacional completa.
