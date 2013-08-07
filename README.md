#ChaordicTests

Desafios de programação propostos pelo Anderson

##/vcards/
Primeiro desafio, constitui em dois programas:
- Um para popular uma planilha do Google com dados obtidos por um formulário em uma página da web
- Outro para transformar as informações desta planilha em arquivos .vcf (vCard)

###Dependências
Para a criação destes programas, utilizei a biblioteca gdata-python-client (Disponível em <http://code.google.com/p/gdata-python-client/>), que fornece acesso a serviços do Google a partir de feeds Atom é necessário tê-la instalada para executar estes programas.
Já para a execução, é necessário ter um interpretador de Python 2.7 (Disponível em <http://python.org/download/>) e no caso do formulário, um servidor HTTP (remoto ou local) com suporte a scripts CGI (ex.: Apache).

###Instalação
- Coloque os arquivos `spreadsheetform.py` e `index.html` em uma pasta de seu servidor que esteja habilitada para executar scripts CGI.
- Abra o endereço do arquivo `index.html` em seu servidor através de um navegador

- Coloque o arquivo `vcard.py` em qualquer pasta
- Inicie uma sessão de console e navegue até o local onde está `vcard.py`
- Digite o comando abaixo, substituindo os parâmetros indicados. (Dependendo da configuração de seu sistema, pode ser necessário usar um caminho absoluto para o interpretador Python)

    `python vcard.py <usuário Google> <senha Google> <planilha> <página> <quantidade de entradas a processar> <caminho de destino>`
