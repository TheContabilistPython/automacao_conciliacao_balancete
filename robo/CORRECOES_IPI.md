# SOLUÇÃO PARA PROBLEMA DE EXTRAÇÃO DO IPI

## Problema Identificado
O script FISCAL_V2.py estava retornando `None` para `total_ipi_recolher` porque:
1. Havia um espaço extra no nome do arquivo HTML
2. A lógica de extração era muito restritiva
3. Não havia múltiplas abordagens para encontrar o valor

## Correções Implementadas

### 1. Correção do Nome do Arquivo
- **Antes**: `"Relatório apuração IPI -   Sequência 22"` (3 espaços)
- **Depois**: `"Relatório apuração IPI -  Sequência 22"` (2 espaços)

### 2. Busca por Múltiplas Variações de Nome
Implementada busca automática por diferentes variações:
- `"Relatório apuração IPI -  Sequência 22"` (2 espaços)
- `"Relatório apuração IPI - Sequência 22"` (1 espaço)
- `"Relatório apuração IPI -   Sequência 22"` (3 espaços)
- Busca por padrão usando `glob` se nenhuma variação exata for encontrada

### 3. Múltiplas Abordagens de Extração
Implementadas 3 abordagens diferentes:

#### Abordagem 1: Busca Específica
- Busca por `<td colspan="7" class="s6">` contendo "Valor do saldo devedor do IPI a recolher"
- Procura valores monetários na mesma linha

#### Abordagem 2: Busca Ampla
- Busca por qualquer célula contendo "IPI a recolher"
- Procura valores monetários na linha correspondente

#### Abordagem 3: Busca por Padrões
- Busca por todos os valores monetários no HTML
- Filtra aqueles que estão próximos a texto relacionado ao IPI

### 4. Melhorias de Debug
- Adicionado debug detalhado que mostra células encontradas
- Verificação de conteúdo HTML (se contém 'IPI' e 'recolher')
- Logs mais informativos para identificar problemas

### 5. Ferramentas de Diagnóstico
Criadas ferramentas auxiliares:
- `TESTE_IPI.bat` - Testa especificamente a extração do IPI
- `ANALISAR_HTML.bat` - Analisa estrutura HTML completa
- `test_ipi_extraction.py` - Script de teste dedicado
- `analyze_html_structure.py` - Analisador de estrutura HTML

### 6. Atualização do Menu Principal
- Adicionadas opções 5 e 6 para diagnóstico
- Menu agora tem 7 opções (incluindo ferramentas de diagnóstico)

### 7. Documentação Atualizada
- README.md atualizado com instruções de diagnóstico
- Explicação das novas ferramentas

## Como Testar a Solução

### Teste Rápido
```bash
# Execute o menu principal
MENU_PRINCIPAL.bat

# Escolha opção 5 para testar IPI
# Escolha opção 6 para análise completa
```

### Teste Específico
```bash
# Teste direto da extração do IPI
TESTE_IPI.bat

# Análise completa da estrutura HTML
ANALISAR_HTML.bat
```

### Teste do Script Principal
```bash
# Execute o script fiscal
FISCAL_V2.bat

# Observe os logs para:
# - "Processando arquivo IPI: [caminho]"
# - "Debug: Encontradas X células contendo 'IPI'"
# - "ipi_recolher (método X): [valor]"
```

## Resultados Esperados
Com essas correções, o script deve:
1. Encontrar o arquivo HTML do IPI mesmo com variações de nome
2. Extrair corretamente o valor 416.231,26 do HTML fornecido
3. Mostrar logs detalhados do processo de extração
4. Não mais retornar `None` para `total_ipi_recolher`

## Arquivos Modificados
- `FISCAL_V2.py` - Script principal com correções
- `MENU_PRINCIPAL.bat` - Menu atualizado com diagnóstico
- `README.md` - Documentação atualizada
- `TESTE_IPI.bat` - Nova ferramenta de teste
- `ANALISAR_HTML.bat` - Nova ferramenta de análise
- `test_ipi_extraction.py` - Script de teste do IPI
- `analyze_html_structure.py` - Analisador de estrutura HTML

## Próximos Passos
1. Execute `TESTE_IPI.bat` para validar a extração
2. Execute `FISCAL_V2.bat` para testar o script completo
3. Verifique se o valor 416.231,26 é extraído corretamente
4. Confirme se não há mais "Verificar" na coluna J para o IPI
