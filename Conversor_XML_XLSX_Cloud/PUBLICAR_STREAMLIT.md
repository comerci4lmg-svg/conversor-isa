# Publicar no Streamlit Community Cloud

1. Crie um repositório no GitHub, por exemplo `conversor-xml-xlsx`.
2. Envie os arquivos deste pacote extraídos para a raiz do repositório (não envie apenas o ZIP).
3. Acesse https://share.streamlit.io/ e conecte sua conta GitHub.
4. Clique em Create app e selecione o repositório.
5. Selecione a branch `main` e informe `streamlit_app.py` em Main file path.
6. Em Advanced settings, selecione Python 3.12 e clique em Deploy.

O requirements.txt instala as dependências automaticamente. Não é preciso configurar secrets.
O iniciar.cmd é apenas para uso local no Windows.
O XML de dados deve ser enviado pelo aplicativo depois de publicado.

O formato SpreadsheetML do arquivo BASE OPER é reconhecido automaticamente.
