# Sistema de Automação Fiscal e Folha de Pagamento

Este projeto contém scripts Python para automação de processos fiscais e de folha de pagamento.

## Estrutura do Projeto

- `robo/` - Scripts principais de automação
  - `FISCAL_V2.py` - Processamento de relatórios fiscais
  - `FOLHA_V3.PY` - Processamento de folha de pagamento
  - `auto_bal_v2.py` - Balancetes automáticos
  - `requirements.txt` - Dependências Python

- `planilhas/` - Planilhas de conciliação
  - `CONCILIACAO_*.xlsx` - Planilhas de conciliação por empresa
  - `balancete/` - Arquivos de balancete

## Pré-requisitos

- Python 3.7+
- Dependências listadas em `requirements.txt`

## Instalação

1. Clone o repositório
2. Instale as dependências:
   ```bash
   pip install -r robo/requirements.txt
   ```

## Uso

### Script Fiscal
```bash
python robo/FISCAL_V2.py
```

### Script de Folha
```bash
python robo/FOLHA_V3.PY
```

## Características

- ✅ Processamento automático de relatórios HTML
- ✅ Integração com planilhas Excel
- ✅ Cálculo automático do último dia do mês
- ✅ Interface gráfica para seleção de empresa e competência
- ✅ Validação de dados com tolerância configurável

## Melhorias Recentes

- Implementação do cálculo dinâmico do último dia do mês
- Correção para meses com 30 dias (junho, abril, setembro, novembro)
- Suporte a anos bissextos para fevereiro

## Contribuição

Para contribuir com o projeto:
1. Faça um fork
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Abra um Pull Request

## Licença

Projeto interno para automação de processos.
