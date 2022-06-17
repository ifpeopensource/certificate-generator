# Certificate Generator

<p align="center">
  <img alt="Contador de linguagens" src="https://img.shields.io/github/languages/count/ifpeopensource/certificate-generator?color=%2304D361">

  <img alt="Tamanho do repositório" src="https://img.shields.io/github/repo-size/ifpeopensource/certificate-generator">

  <a href="https://ifpeopensource.com.br">
    <img src="https://img.shields.io/badge/IFPE Open Source-0a0a0a?logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNTUgMjUyIj48ZGVmcz48c3R5bGU+LmNscy0xe2ZpbGw6IzJmOWU0MTt9LmNscy0ye2ZpbGw6I2M4MTkxZTt9PC9zdHlsZT48L2RlZnM+PHRpdGxlPkJyYW5kIElGUEUgT3BlbiBTb3VyY2U8L3RpdGxlPjxnIGlkPSJMb2dvbWFyayI+PGcgaWQ9IkxvZ29tYXJrLTIiIGRhdGEtbmFtZT0iTG9nb21hcmsiPjxyZWN0IGNsYXNzPSJjbHMtMSIgeD0iMTQyLjMiIHk9IjEwLjUyIiB3aWR0aD0iMTA3LjIiIGhlaWdodD0iMTA3LjIiIHJ4PSI5Ii8+PHJlY3QgY2xhc3M9ImNscy0xIiB4PSIxNDIuMyIgeT0iMTM5LjE2IiB3aWR0aD0iMTA3LjIiIGhlaWdodD0iMTA3LjIiIHJ4PSI5Ii8+PHJlY3QgY2xhc3M9ImNscy0xIiB4PSIxMy42NiIgeT0iMTM5LjE2IiB3aWR0aD0iMTA3LjIiIGhlaWdodD0iMTA3LjIiIHJ4PSI5Ii8+PGcgaWQ9Ik9wZW5fU291cmNlX1N5bWJvbCIgZGF0YS1uYW1lPSJPcGVuIFNvdXJjZSBTeW1ib2wiPjxwYXRoIGNsYXNzPSJjbHMtMiIgZD0iTTY2LjczLDUuNjRBNjAuODUsNjAuODUsMCwwLDAsNDQuOSwxMjMuM2EzLjc0LDMuNzQsMCwwLDAsMy0uMTMsMy44NSwzLjg1LDAsMCwwLDItMi4yOUw2MS4zMSw4NC40MkEzLjgxLDMuODEsMCwwLDAsNTkuNTIsODAsMTUuMjMsMTUuMjMsMCwxLDEsNzQsODBhMy44LDMuOCwwLDAsMC0xLjc5LDQuNDdsMTEuNDIsMzYuNDZhMy44NCwzLjg0LDAsMCwwLDIsMi4zLDMuOTEsMy45MSwwLDAsMCwxLjY2LjM4LDQsNCwwLDAsMCwxLjM2LS4yNUE2MC44NSw2MC44NSwwLDAsMCw2Ni43Myw1LjY0WiIvPjwvZz48L2c+PC9nPjwvc3ZnPg=="></img>
  </a>

  <a href="https://github.com/ifpeopensource/certificate-generator/commits/master">
    <img alt="Último commit" src="https://img.shields.io/github/last-commit/ifpeopensource/certificate-generator">
  </a>

  <a href="https://github.com/ifpeopensource/certificate-generator/issues">
    <img alt="Issues" src="https://img.shields.io/github/issues/ifpeopensource/certificate-generator">
  </a>
  <a href="https://github.com/ifpeopensource/certificate-generator/blob/master/LICENSE" target="_blank">
    <img alt="License" src="https://img.shields.io/badge/licença-MIT-brightgreen"/>
  </a>
</p>

Programa em Python para automatizar a geração de certificados PDF a partir de um modelo em PPTX (arquivo do PowerPoint).

## Dependências

Para o funcionamento do programa, é necessário ter instalado as seguintes dependências:

- [LibreOffice](https://www.libreoffice.org/)
- [Python 3.7+](https://www.python.org/)

### Como utilizar

  1. Instale as dependências
      - Linux - APT

          1. Instale o LibreOffice

              ```console
              sudo apt update
              sudo apt install libreoffice
              ```

          2. Crie um virtual environment para o Python 3.7+ e instale as dependências do programa

              ```console
              python3 -m venv venv
              source venv/bin/activate
              pip install -r requirements.txt
              ```

  2. Crie um arquivo `names.txt` com os nomes das pessoas separados por linha.
  3. Crie um arquivo `model.pptx` com o slide modelo do certificado.
     - Os campos com o texto `{{name}}` serão substituídos pelos nomes das pessoas.
  4. Rode o programa.

      ```console
      python main.py ./output_file.pdf
      ```

  5. Os certificados serão gerados no diretório `./output`. Caso o parâmetro seja informado, um arquivo PDF contendo todos os certificados será gerado.

## Time

Este projeto é mantido por esses [incríveis contribuidores](https://github.com/ifpeopensource/certificate-generator/graphs/contributors).

## 🤝 Contribuir

Contribuições, issues e pedidos de features são bem-vindos!<br />Sinta-se livre para checar a [página de issues](https://github.com/ifpeopensource/certificate-generator/issues).

- Crie um fork;
- Crie um branch com a sua feature: `git checkout -b my-feature`;
- Faça um commit com as mudanças: `git commit -m 'feat: My new feature'`;
- Faça um push para o seu branch: `git push origin my-feature`.

Após a sua pull request ser aceita, você pode excluir o seu branch.

## Demonstre o seu apoio

Dê uma ⭐️ se este projeto lhe ajudou!

## 📝 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

***
Feito com ♥ no [<img src="https://github.com/ifpeopensource.png" width="15px"/> IFPE Open Source](https://ifpeopensource.com.br)

## English

Program in Python to automate the generation of PDF certificates from a PPTX model (PowerPoint file).

### Dependencies

For the program to work, it is necessary to have installed the following dependencies:

- [LibreOffice](https://www.libreoffice.org/)
- [Python 3.7+](https://www.python.org/)

### Usage

  1. Install dependencies
      - Linux - APT

          1. Install LibreOffice

              ```console
              sudo apt update
              sudo apt install libreoffice
              ```

          2. Create virtual environment and install dependencies

              ```console
              python3 -m venv venv
              source venv/bin/activate
              pip install -r requirements.txt
              ```

  2. Create a `names.txt` file with the names of the people you want to display;
  3. Create a `model.pptx` file with the certificate model slide;
     - The fields with `{{name}}` will be replaced with the person name;
  4. Run the program

      ```console
      python main.py ./output_file.pdf
      ```

  5. The output files will be generated in the `./output` directory. If the parameter is informed, a PDF file containing all the certificates will be generated.

### Team

This project is maintained by these [incredible contributors](https://github.com/ifpeopensource/certificate-generator/graphs/contributors).

### 🤝 Contributing

Contributions, issues and feature requests are welcome!<br />Feel free to check the [issues page](https://github.com/ifpeopensource/certificate-generator/issues).

- Create a fork;
- Create a branch with your feature: `git checkout -b my-feature`;
- Commit changes: `git commit -m 'feat: My new feature'`;
- Push to your branch: `git push origin my-feature`.

After your pull request is accepted, you can delete your branch.

### Show your support

Give a ⭐️ if this project helped you!

### 📝 License

This project is licensed under the MIT license. See the file [LICENSE](LICENSE) for more details.

***
Made with ♥ at [<img src="https://github.com/ifpeopensource.png" width="15px"/> IFPE Open Source](https://ifpeopensource.com.br)
