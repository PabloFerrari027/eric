# Projeto Eric

Este projeto automatiza a busca e atualização de dados de SKU e DROP através da API do Asana, gerando uma planilha Excel com os resultados.

## 📋 Funcionalidades

- ✅ Busca automática de dados na API do Asana
- ✅ Geração/atualização de planilha Excel
- ✅ Detecção automática de novos SKUs
- ✅ Fechamento automático de arquivos Excel em uso
- ✅ Abertura automática da planilha após processamento
- ✅ Suporte multiplataforma (Windows, macOS, Linux)

## 🛠️ Pré-requisitos

Antes de usar o projeto, certifique-se de ter instalado:

- **Node.js** (versão 16 ou superior) - [Download aqui](https://nodejs.org/)
- **Excel ou LibreOffice Calc** - Para visualizar as planilhas
- **Git** (opcional) - Para clonar o repositório

## 🚀 Instalação

### 1. Clone ou baixe o projeto
```bash
gh repo clone PabloFerrari027/eric
cd eric
```

### 2. Instale as dependências
```bash
npm install
```

## ⚙️ Configuração

### Token da API do Asana
O projeto já vem configurado com um token da API do Asana. Se precisar alterar:

1. Abra o arquivo `src/fetch-data.ts`
2. Localize a linha com `Authorization`
3. Substitua pelo seu token

### Lista de SKUs
Para modificar os SKUs que serão processados:

1. Abra o arquivo `src/index.ts`
2. Localize o array `skus`
3. Adicione, remova ou modifique os SKUs conforme necessário

```typescript
const skus: Array<string> = [
  "TS751.01",
  "TS752.01",
  // Adicione mais SKUs aqui
];
```

## 🎯 Como usar

### Executar o projeto
```bash
npm start
```

### O que acontece quando você executa:

1. **Verificação de arquivo**: O sistema verifica se existe uma planilha `SKU_DROP.xlsx`
2. **Leitura de dados**: Se a planilha existir, lê os dados já processados
3. **Busca na API**: Consulta a API do Asana para novos SKUs
4. **Atualização**: Adiciona apenas os novos dados à planilha
5. **Abertura automática**: Abre a planilha no Excel/Calc

### Exemplo de saída no terminal:
```
🚀 Iniciando processamento...
✅ Arquivo SKU_DROP.xlsx disponível para uso.
📊 15 registros existentes encontrados.
🔍 Buscando 3 novos SKUs...
🔍 Buscando: TS759.01
  ➕ Adicionando: TS759.01-P - DROP: DROP 1
  ➕ Adicionando: TS759.01-M - DROP: DROP 1
💾 Atualizando planilha...
✅ Planilha atualizada!
📊 Novos registros adicionados: 2
📋 Total de registros: 17
📂 Abrindo arquivo: SKU_DROP.xlsx
✅ Arquivo Excel aberto com sucesso!
```

## 📁 Estrutura do projeto

```
eric/
├── src/
│   ├── index.ts                    # Arquivo principal
│   ├── fetch-data.ts              # Busca dados na API
│   ├── read-data.ts               # Lê dados da planilha
│   ├── write-data.ts              # Escreve dados na planilha
│   ├── ensure-file-access.ts      # Garante acesso ao arquivo
│   ├── close-specific-excel-file.ts # Fecha arquivos específicos
│   ├── is-file-in-use.ts          # Verifica se arquivo está em uso
│   ├── open-excel-file.ts         # Abre arquivo no Excel
│   └── types.ts                   # Definições de tipos
├── package.json                   # Configurações do projeto
├── tsconfig.json                  # Configurações do TypeScript
├── .gitignore                     # Arquivos ignorados pelo Git
└── SKU_DROP.xlsx                  # Planilha gerada (após primeira execução)
```

## 📊 Formato da planilha

A planilha gerada terá duas colunas:

| SKU        | DROP   |
|------------|--------|
| TS751.01-P | DROP 1 |
| TS751.01-M | DROP 1 |
| TS751.01-G | DROP 1 |

## ⚠️ Solução de problemas

### Arquivo Excel está aberto/em uso
Se você receber o erro "arquivo está em uso":
1. Feche manualmente o arquivo Excel
2. Execute o comando novamente
3. O sistema tentará fechar automaticamente na próxima execução

### Erro de permissão no Windows
Execute o terminal como administrador:
1. Clique com botão direito no Command Prompt ou PowerShell
2. Selecione "Executar como administrador"
3. Execute `npm start`

### Node.js não encontrado
Certifique-se de que o Node.js está instalado corretamente:
```bash
node --version
npm --version
```

### Erro de dependências
Reinstale as dependências:
```bash
rm -rf node_modules
rm package-lock.json
npm install
```

## 🔧 Personalização

### Alterar nome do arquivo Excel
No arquivo `src/index.ts`, modifique:
```typescript
const filePath = "./MEU_ARQUIVO.xlsx";
```

### Adicionar novas colunas
1. Modifique o tipo `Row` em `src/types.ts`
2. Ajuste as funções de leitura e escrita conforme necessário

### Alterar campos da API
No arquivo `src/fetch-data.ts`, modifique os campos consultados conforme a estrutura da sua API do Asana.

## 📞 Suporte

Se encontrar problemas:
1. Verifique se todas as dependências estão instaladas
2. Confirme se o Node.js está na versão correta
3. Verifique se o arquivo Excel não está aberto em outro programa
4. Execute com permissões de administrador se necessário

## 🚀 Scripts disponíveis

- `npm start` - Executa o projeto completo
- `npm install` - Instala dependências

## 📝 Notas importantes

- O projeto detecta automaticamente se já existem dados na planilha
- Apenas novos SKUs são processados a cada execução
- O arquivo Excel é aberto automaticamente após o processamento
- O sistema funciona em Windows, macOS e Linux
- Os dados são salvos na aba "SKU x DROP"
