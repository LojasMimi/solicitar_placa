
# 🏷️ Solicitação de Placas

Aplicação desenvolvida em **Python com Streamlit**, para uso interno e para facilitar a solicitação de placas de produtos, permitindo consultas individuais, processamento em lote e geração automática de um formulário Excel padronizado.

---

## 🚀 Funcionalidades

- 📌 Solicitação individual de produtos via código de barras  
- 📦 Solicitação em lote através de arquivo Excel  
- 🖼️ Visualização do tamanho da placa por imagens  
- 📊 Relatório com produtos solicitados  
- 🗑️ Remoção de produtos da solicitação  
- 📥 Geração e download automático do formulário 'solicitar placa.xlsx'

---

## 🛠️ Tecnologias utilizadas

- Python 3
- Streamlit
- Requests
- Pandas
- OpenPyXL

---

## 📁 Estrutura do projeto

```

├── app_solicitar_placa.py
├── solicitar placa.xlsx
├── imagens/
│   ├── HORIZONTAL.png
│   ├── VERTICAL.png
│   ├── MEIA_FOLHA.jpg
│   ├── UM_QUARTO_FOLHA.jpg
│   └── ETIQUETA_GONDOLA.jfif
└── README.md

````

---

## ▶️ Como executar o projeto

1. Clone o repositório:
```bash
git clone https://github.com/LojasMimi/solicitar_placa
````

2. Instale as dependências:

```bash
pip install -r requirements.txt
```

3. Execute a aplicação:

```bash
streamlit run app_solicitar_placa.py
```

---

## 📝 Observações

* A aplicação depende do arquivo **`solicitar placa.xlsx`** como modelo base.
* A pasta **`imagens/`** é obrigatória para exibição correta dos tamanhos das placas.
* É necessário acesso à API do Varejo Fácil para consulta de produtos.

---

## 📌 Status do projeto

✅ Funcional

📦 Pronto para uso interno

🔧 Manutenções e melhorias futuras podem ser adicionadas

---

Desenvolvido para otimizar o processo de solicitação de placas de forma simples e eficiente por Pablo Dantas.

