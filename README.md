## Tratamento de dados com java
Esse projeto visava tratar dados de uma planilha de uma biblioteca, corrigindo seus dados e buscando os que estavam incompletos via requisição HTTP à API do Google Books. Em seguida, foi utilizado o banco de dados MySQL para armazenamento dos dados da planilha. 

## Tecnologias utilizadas
- **Aspose Cells**: Biblioteca para manipular arquivos de planilha
- **Gson**: Biblioteca que serializa objetos Java para JSON (e vice-versa)
- **MySQL**: Banco de dados para armazenamento dos dados requisitados
- **Http Client Java**: Biblioteca Java nativa para requisições http

### Para executar, baixe o projeto, troque as variáveis de conexão e os caminhos de arquivos locais e execute o seguinte comando:
```maven
mvn exec:java
```
