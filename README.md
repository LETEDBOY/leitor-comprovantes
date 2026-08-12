# Leitor de Comprovantes

App Streamlit que lê fotos de comprovantes (posto, lavanderia, etc.) usando a API da Claude
e preenche a planilha de prestação de contas.

## Deploy no Streamlit Community Cloud (pra usar do iPhone)

1. Suba este repositório no GitHub (pode ser privado).
2. Entre em https://share.streamlit.io, conecte sua conta GitHub.
3. Clique em "New app", escolha este repo e o arquivo `leitor_comprovantes.py`.
4. Em "Advanced settings" > "Secrets", cole:
   ```
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
5. Deploy. Você recebe uma URL tipo `https://seu-app.streamlit.app`.
6. Abra essa URL no Safari do iPhone. Em "Compartilhar" > "Adicionar à Tela de Início" pra virar um atalho tipo app.

No campo de upload de foto, o Safari no iPhone abre automaticamente a opção de
"Tirar Foto" ou escolher da galeria.
