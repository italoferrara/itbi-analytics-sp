# ITBI Analytics SP — Web v0.1

Versão Streamlit da ferramenta de consulta de transações imobiliárias com ITBI da Prefeitura de São Paulo.

## Executar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar no Streamlit Community Cloud

1. Envie estes arquivos para um repositório GitHub.
2. Entre em https://share.streamlit.io com a conta GitHub.
3. Clique em **Create app** / **Deploy an app**.
4. Selecione o repositório e `app.py` como entrypoint.
5. Deploy.
6. Na primeira abertura do site, abra o menu lateral e clique **Atualizar base da Prefeitura** para criar a base no servidor.

> Esta v0.1 usa armazenamento local do servidor Streamlit para a base. É adequada para piloto. Em uma versão permanente/comercial, mover a base para PostgreSQL/Supabase evita reconstruções quando o container for reiniciado.
