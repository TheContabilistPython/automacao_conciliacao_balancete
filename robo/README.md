# 📂 SISTEMA DE CONCILIAÇÃO FISCAL E FOLHA

## 🚀 INSTALAÇÃO EM NOVA MÁQUINA

### 1. PRÉ-REQUISITOS
- **Python 3.7 ou superior** instalado
  - Download: https://python.org/downloads/
  - ✅ Marcar "Add Python to PATH" durante instalação
- **Windows 10/11**

### 2. ESTRUTURA DE PASTAS NECESSÁRIA
Crie as seguintes pastas na máquina de destino:

```
C:\
├── projeto\
│   ├── robo\                    # Scripts Python e .bat
│   ├── planilhas\               # Planilhas Excel de saída
│   └── empresas.csv             # Arquivo com códigos/nomes das empresas
├── relatorios_fiscal\           # Relatórios HTML fiscais
└── relatorios_folha\           # Relatórios HTML de folha
```

### 3. INSTALAÇÃO PASSO A PASSO

#### Passo 1: Copiar arquivos
1. Copie toda a pasta `C:\projeto\robo\` para a nova máquina
2. Copie o arquivo `C:\projeto\empresas.csv`
3. Crie as pastas de relatórios se não existirem

#### Passo 2: Instalar dependências
1. Abra o prompt de comando como **Administrador**
2. Navegue até: `cd C:\projeto\robo\`
3. Execute: `INSTALAR_DEPENDENCIAS.bat`

#### Passo 3: Verificar instalação
Execute o teste: `python -c "import bs4, openpyxl, pyautogui; print('OK - Todas as bibliotecas estão instaladas!')"`

### 4. ARQUIVOS INCLUÍDOS

#### Scripts Python:
- `FISCAL_V2.py` - Conciliação fiscal
- `FOLHA_V3.PY` - Conciliação folha de pagamento  
- `auto_bal_v2.py` - Processamento balancete

#### Scripts .bat:
- `MENU_PRINCIPAL.bat` - Menu principal para executar todos os scripts
- `FISCAL_V2.bat` - Executa apenas conciliação fiscal
- `FOLHA_V3.bat` - Executa apenas conciliação folha
- `auto_bal_v2.bat` - Executa apenas balancete
- `TESTE_IPI.bat` - Diagnóstico de extração do IPI
- `INSTALAR_DEPENDENCIAS.bat` - Instala dependências Python

#### Arquivos de configuração:
- `requirements.txt` - Lista de dependências Python
- `empresas.csv` - Códigos e nomes das empresas

### 5. COMO USAR

#### Opção 1: Menu principal (Recomendado)
1. Execute: `MENU_PRINCIPAL.bat`
2. Escolha a opção desejada (1-5)

#### Opção 2: Scripts individuais
- Para balancete: `auto_bal_v2.bat`
- Para fiscal: `FISCAL_V2.bat` 
- Para folha: `FOLHA_V3.bat`

### 6. ESTRUTURA DO ARQUIVO empresas.csv
```csv
codigo;nome
1001;EMPRESA EXEMPLO LTDA
1002;OUTRA EMPRESA S/A
```

### 7. NOMENCLATURA DOS ARQUIVOS HTML
Os relatórios HTML devem seguir o padrão:
```
Empresa [CODIGO] - [NOME] - [TIPO_RELATORIO] - Sequência 22 - Ordem [X] 01-[MM-AAAA] 31-[MM-AAAA].htm
```

### 8. DIAGNÓSTICO E RESOLUÇÃO DE PROBLEMAS

#### 8.1 Teste de Extração do IPI
Se o valor do IPI não estiver sendo extraído corretamente, execute:
```
TESTE_IPI.bat
```

Este script irá:
- Verificar se os arquivos HTML do IPI existem
- Testar diferentes variações de nomes de arquivo
- Mostrar quais valores foram encontrados
- Ajudar a identificar problemas de estrutura HTML

#### 8.2 Problemas Comuns

#### Erro: "Python não encontrado"
- Reinstale Python marcando "Add to PATH"
- Ou adicione Python manualmente ao PATH do Windows

#### Erro: "Módulo não encontrado"
- Execute: `pip install -r requirements.txt`
- Ou: `INSTALAR_DEPENDENCIAS.bat`

#### Erro: "Arquivo não encontrado"
- Verifique se a estrutura de pastas está correta
- Verifique se os relatórios HTML estão nos locais corretos

#### Erro: "Permissão negada"
- Execute como Administrador
- Verifique se os arquivos Excel não estão abertos

### 9. REQUISITOS DO SISTEMA
- **OS**: Windows 10/11
- **RAM**: Mínimo 4GB
- **Espaço**: ~1GB para Python + dependências
- **Processador**: Qualquer processador moderno

### 10. CONTATO E SUPORTE
Para problemas técnicos, verifique:
1. Logs de erro no console
2. Estrutura de pastas
3. Nomenclatura dos arquivos
4. Permissões de acesso

---
**Versão**: 2.0  
**Última atualização**: Janeiro 2025
