# Word AI Styler 📝

Um processador inteligente de documentos Word que utiliza IA (GPT-4.1) para identificar e aplicar estilos automaticamente em documentos educacionais, mantendo 100% do conteúdo original.

## 🚀 Funcionalidades

### Principais
- **Identificação Inteligente**: Usa GPT-4.1 para identificar automaticamente diferentes elementos do documento (questões, alternativas, gabaritos, títulos, etc.)
- **Aplicação de Estilos Customizados**: Define e aplica estilos personalizados do Word com cores e formatações específicas
- **Preservação Total**: Mantém 100% do conteúdo original, incluindo imagens, fórmulas matemáticas, tabelas e formatações
- **Interface Visual**: Interface web intuitiva para configurar estilos e acompanhar o processamento
- **Templates Salvos**: Salve e reutilize configurações de estilos

### Recursos Avançados
- **Dupla Verificação**: Sistema de duas passadas para garantir máxima precisão na identificação
- **Análise Contextual**: Segunda passada analisa elementos não marcados considerando o contexto
- **Remoção Seletiva**: Opção para remover seções específicas (como capas ou cartões resposta)
- **Suporte a Imagens Inline**: Mantém estilos mesmo em parágrafos com imagens
- **Estatísticas Detalhadas**: Relatório completo do processamento

## 🛠️ Tecnologias

### Backend
- **Python 3.x**
- **Flask** - API REST
- **python-docx** - Manipulação de documentos Word
- **OpenAI API** - Integração com GPT-4.1
- **Flask-CORS** - Suporte para requisições cross-origin

### Frontend
- **React** - Interface de usuário
- **Tailwind CSS** - Estilização
- **Lucide Icons** - Ícones
- **JavaScript ES6+**

## 📋 Pré-requisitos

- Python 3.8 ou superior
- Node.js (para servir o frontend)
- Chave de API da OpenAI com acesso ao GPT-4.1
- Sistema operacional: Windows, macOS ou Linux

## 🔧 Instalação

### 1. Clone o repositório
```bash
git clone https://github.com/seu-usuario/word-ai-styler.git
cd word-ai-styler
```

### 2. Configure o ambiente Python
```bash
# Crie um ambiente virtual
python -m venv venv

# Ative o ambiente virtual
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

# Instale as dependências
pip install -r requirements.txt
```

### 3. Configure as variáveis de ambiente
Crie um arquivo `.env` na raiz do projeto:
```env
OPENAI_API_KEY=sua_chave_api_aqui
FLASK_ENV=development
FLASK_PORT=5000
```

### 4. Estrutura de pastas
O sistema criará automaticamente as seguintes pastas:
- `uploads/` - Arquivos enviados temporariamente
- `output/` - Documentos processados
- `temp/` - Arquivos temporários

## 🚀 Uso

### 1. Inicie o servidor backend
```bash
python run.py
```
O servidor estará disponível em `http://localhost:5000`

### 2. Abra o frontend
Abra o arquivo `frontend/index.html` em seu navegador ou sirva com um servidor HTTP:
```bash
# Com Python
python -m http.server 8000 --directory frontend

# Ou com Node.js
npx http-server frontend -p 8000
```

### 3. Configure os estilos
1. **Defina os estilos**: Configure nome, estilo do Word, cor e prompt de identificação
2. **Adicione marcadores**: Cada estilo tem um marcador único (ex: `[[ENUNCIADO]]`)
3. **Configure prompts**: Descreva como a IA deve identificar cada elemento

### 4. Processe o documento
1. Faça upload do arquivo `.docx`
2. Insira sua chave API da OpenAI
3. Defina um nome para o projeto
4. Clique em "Processar Documento"

## 📁 Estrutura do Projeto

```
word-ai-styler/
├── api/
│   ├── __init__.py
│   └── routes.py          # Endpoints da API
├── backend/
│   ├── __init__.py
│   ├── ai_processor.py    # Integração com GPT-4.1
│   ├── config.py          # Configurações
│   ├── document_reader.py # Leitura de documentos
│   ├── style_applier.py  # Aplicação de estilos
│   ├── document_splitter.py # Divisão de documentos
│   └── file_manager.py    # Gerenciamento de arquivos
├── frontend/
│   └── index.html         # Interface web
├── uploads/               # Arquivos temporários
├── output/                # Documentos processados
├── temp/                  # Arquivos temporários
├── .env                   # Variáveis de ambiente
├── .gitignore
├── requirements.txt       # Dependências Python
├── run.py                 # Arquivo principal
└── README.md
```

## 🎯 Casos de Uso

### Ideal para:
- **Editoras Educacionais**: Padronização de simulados e materiais didáticos
- **Professores**: Formatação automática de provas e exercícios
- **Instituições de Ensino**: Criação de material padronizado
- **Preparatórios**: Organização de questões e gabaritos

### Tipos de documento suportados:
- Simulados e provas
- Listas de exercícios
- Material didático com questões
- Documentos com padrões repetitivos

## ⚙️ Configuração Avançada

### Estilos Personalizados
Cada estilo pode ter:
- **Nome**: Identificação do estilo
- **Estilo Word**: Nome do estilo no documento final
- **Cor**: Cor de destaque (formato hex)
- **Marcador**: Tag única para identificação
- **Prompt**: Instrução para a IA identificar o elemento
- **Permite Imagens Inline**: Se o estilo pode ser aplicado em parágrafos com imagens

### Exemplo de Configuração
```javascript
{
    name: "Enunciado",
    wordStyle: "Questão",
    marker: "[[ENUNCIADO]]",
    prompt: "Identifique enunciados que começam com números seguidos de ponto ou parêntese",
    color: "#dc2626",
    allowInlineImages: true
}
```

## 🔍 Solução de Problemas

### Documento sem estilos
- Verifique se os prompts estão claros e específicos
- Aumente o detalhamento dos prompts de identificação
- Verifique os logs do console para erros

### Imagens não aparecem
- O sistema preserva todas as imagens automaticamente
- Certifique-se de que o documento original não está corrompido

### Taxa de marcação baixa
- O sistema faz automaticamente uma segunda passada
- Ajuste os prompts para serem mais específicos
- Verifique se os marcadores estão corretos

## 📈 Limitações

- Arquivos Word devem estar no formato `.docx`
- Tamanho máximo de arquivo: 50MB
- Requer conexão com internet para usar a API da OpenAI
- Processamento pode levar alguns minutos dependendo do tamanho

