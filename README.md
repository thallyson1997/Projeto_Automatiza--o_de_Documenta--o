# 📄 Projeto Automatização de Documentação

Uma aplicação web para **gerar automaticamente documentos Word** com inserção de imagens em múltiplas páginas. Ideal para automatizar a criação de relatórios, documentações técnicas e formulários com layouts consistentes.

🌐 **[Acesse o site aqui!](https://projeto-automatiza-documento.onrender.com/)**

---

## ✨ Funcionalidades

- ✅ **Geração de documentos Word (.docx)** com múltiplas páginas
- ✅ **Inserção de 1 a 4 imagens por página** com layouts automáticos
- ✅ **Múltiplos formulários** em um único documento
- ✅ **Sincronização de campos** entre páginas (repetir unidade e data)
- ✅ **Notificações visuais** e responsivas
- ✅ **Download automático** dos documentos gerados
- ✅ **Interface intuitiva** e amigável

### Layouts de Imagens Suportados:
- **1 imagem**: Centralizada
- **2 imagens**: Empilhadas verticalmente (com espaçamento)
- **3 imagens**: 2 lado a lado + 1 abaixo
- **4 imagens**: Grid 2×2

---

## 🛠️ Tecnologias

### Backend
- **Python 3.13+**
- **Flask** - Framework web
- **python-docx** - Manipulação de documentos Word
- **lxml** - Processamento de XML
- **Gunicorn** - Servidor WSGI para produção

### Frontend
- **HTML5** - Estrutura
- **CSS3** - Estilos responsivos
- **JavaScript Vanilla** - Interatividade sem dependências

### Deployment
- **Render.com** - Hosting em nuvem

---

## 📦 Instalação Local

### Pré-requisitos
- Python 3.13 ou superior
- pip (gerenciador de pacotes Python)
- Git

### Passos

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/thallyson1997/Projeto_Automatiza--o_de_Documenta--o
   cd Projeto_Automatiza--o_de_Documenta--o
   ```

2. **Crie um ambiente virtual:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/Mac
   # ou
   venv\Scripts\activate  # Windows
   ```

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Execute a aplicação:**
   ```bash
   python main.py
   ```

5. **Acesse no navegador:**
   ```
   http://localhost:5000
   ```

---

## 🚀 Deployment (Render.com)

### Configuração

O projeto inclui arquivos de configuração para deployment automático:

- **`Procfile`** - Define o comando de inicialização
- **`render.yaml`** - Configuração específica do Render
- **`requirements.txt`** - Dependências Python

### Deploy Automático

1. Faça push para o GitHub:
   ```bash
   git add .
   git commit -m "Deploy configuration"
   git push
   ```

2. No [Render Dashboard](https://dashboard.render.com):
   - Crie um novo **Web Service**
   - Conecte seu repositório GitHub
   - Defina o **Start Command**: `gunicorn main:app --bind 0.0.0.0:10000`
   - O deploy será automaticamente acionado a cada push

---

## 📁 Estrutura do Projeto

```
Projeto_Automatização_de_Documentação/
├── main.py                          # Aplicação Flask principal
├── requirements.txt                 # Dependências Python
├── Procfile                         # Configuração Render/Heroku
├── render.yaml                      # Configuração específica Render
├── .gitignore                       # Arquivos ignorados pelo Git
│
├── functions/
│   ├── __init__.py
│   └── document_generator.py        # Gerador de documentos Word
│
├── templates/
│   ├── index.html                   # Página inicial
│   └── upload.html                  # Página de upload/geração
│
├── static/
│   ├── css/
│   │   └── style.css                # Estilos CSS
│   └── js/
│       └── script.js                # JavaScript frontend
│
└── documento/                       # Arquivos temporários
    └── modelo.docx                  # Template base do documento
```

---

## 💡 Como Usar

### 1. **Página Inicial**
   - Visualize a descrição do projeto
   - Acesse o formulário de geração

### 2. **Gerar Documento**
   - Preencha os campos obrigatórios:
     - **Unidade**: Nome da unidade/departamento
     - **Data**: Data do documento
     - **Legenda**: Descrição/título
     - **Imagens**: Selecione 1 a 4 imagens

### 3. **Recursos Especiais**
   - **Repetir em todos**: Marca para sincronizar unidade/data entre páginas
   - **+ Novo Formulário**: Adicione múltiplas páginas
   - **Gerar Documento**: Baixe o arquivo .docx gerado

### 4. **Download**
   - O documento é automaticamente baixado como `documentacao.docx`

---

## 🔧 Variáveis de Ambiente

Defina estas variáveis para personalização:

```bash
FLASK_ENV=production          # Ambiente de produção
FLASK_APP=main.py             # Arquivo principal
PYTHONUNBUFFERED=1            # Logs em tempo real
```

---

## 📊 Exemplos de Uso

### Relatório com 1 página, 4 imagens:
```
Unidade: Engenharia Civil
Data: 2026-02-03
Legenda: Inspeção de Obra
Imagens: 4 fotos do local
```
**Resultado**: Document com 3 imagens modelo + 4 imagens do usuário

### Relatório com 3 páginas, 2 imagens cada:
```
Formulário 1: Unidade A, Data 2026-02-03, 2 imagens
Formulário 2: Unidade B, Data 2026-02-04, 2 imagens
Formulário 3: Unidade C, Data 2026-02-05, 2 imagens
```
**Resultado**: Documento com 3 páginas, 9 imagens totais

---

## 🐛 Troubleshooting

### "Imagens não aparecem no Word"
- Certifique-se de que os arquivos estão em formato compatível (JPG, PNG)
- Verifique o tamanho máximo das imagens (recomendado < 5MB cada)

### "Erro 404 ao acessar o site"
- Verifique se o Start Command está configurado corretamente no Render
- Limpe o cache do navegador (Ctrl+Shift+Delete)

### "Aplicação lenta ao gerar documento"
- Reduza o tamanho das imagens
- Processe em segundo plano (em desenvolvimento)

---

## 📝 Changelog

### v1.0.0 (Fevereiro 2026)
- ✅ Geração de documentos com múltiplas páginas
- ✅ Suporte a 1-4 imagens por página
- ✅ Sincronização de campos
- ✅ Notificações visuais
- ✅ Deploy em produção

---

## 👨‍💼 Autor

**Thallyson Fontenelle**  
- GitHub: [@thallyson1997](https://github.com/thallyson1997)
- Email: thallyson.gabriel@discente.ufma.br

---

## 📄 Licença

Este projeto está sob licença MIT. Veja o arquivo LICENSE para detalhes.

---

## 🤝 Contribuições

Contribuições são bem-vindas! Para contribuir:

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

---

## 📞 Suporte

Se encontrar problemas ou tiver dúvidas:

1. Verifique os [Logs no Render Dashboard](https://dashboard.render.com)
2. Consulte a seção [Troubleshooting](#-troubleshooting)
3. Abra uma [Issue no GitHub](https://github.com/thallyson1997/Projeto_Automatiza--o_de_Documenta--o/issues)

---

**Desenvolvido com ❤️ em Python**
