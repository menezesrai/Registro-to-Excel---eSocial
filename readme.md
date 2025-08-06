# Extrator de Dados de PDFs para Excel

Olá! Me chamo Raí Menezes e este é mais um dos meus primeiros projetos com Python. Essa aplicação surgiu pois outro afazer do meu dia-a-dia era
criar planilhas do Excel para lojas especificas com os dados de funcionários, este programa me permitiu automatizar essa área completamente, precisando apenas
das fichas dos funcionários (como a encontrada na pasta _data_)

O código foi feito para funcionar especificamente com arquivos como os que estão na pasta _data_. Porém, é simples adaptar isso para funcionar com outros diretórios — basta alterar os campos que o programa procura ao ler o PDF (_começa na linha 33_), não esquecendo também de alterar os campos, seja na função que cria a planilha do excel como na que puxa os dados do PDF.

---

### O programa conta com:

- 📑 **Leitura inteligente de PDFs** – Usa a biblioteca `pdfplumber` para buscar e extrair automaticamente campos relevantes de documentos escaneados;
- 📊 **Exportação organizada para Excel** – Cria uma planilha formatada com as informações extraídas, já com títulos e estilos aplicados com `openpyxl`;

### Instalação

Não se esqueça de rodar:

pip install -r requirements.txt


Grande abraço,  
**Raí Menezes**
