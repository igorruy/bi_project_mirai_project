# Mirai BI Reports

Solução para transformar o XML exportado do Microsoft Project em dois BIs HTML do Projeto Mirai:

- `report_gerencial_mirai.html` — visão gerencial, Curva S, SPI, EVM e indicadores por frente.
- `report_atrasos_mirai.html` — visão de atrasos, desvios BL5/BL0 e responsáveis.

O projeto mantém o uso por linha de comando em `mirai_reports.py` e adiciona uma interface web gratuita com Streamlit em `app.py`, permitindo upload do XML, visualização dos indicadores em tela e download automático dos dois HTMLs.

## Executar localmente com Streamlit

```bash
pip install -r requirements.txt
streamlit run app.py
```

Na tela aberta pelo Streamlit:

1. Faça upload do XML do Microsoft Project na barra lateral.
2. Opcionalmente informe a data de status; se não informar, a aplicação lê o `StatusDate` do XML.
3. Aguarde o processamento.
4. Visualize os BIs em abas e baixe os HTMLs gerados.

## Publicar gratuitamente no Streamlit Community Cloud

1. Suba este repositório para o GitHub.
2. Acesse [share.streamlit.io](https://share.streamlit.io/).
3. Crie um novo app apontando para este repositório.
4. Informe `app.py` como arquivo principal.
5. Publique o app.

O arquivo `.streamlit/config.toml` define `maxUploadSize = 200`, adequado para o XML atual de aproximadamente 163 MB. Se o XML crescer acima desse limite, será necessário reduzir o arquivo, dividir a entrada ou avaliar outro serviço de hospedagem.

## Executar por linha de comando

```bash
python mirai_reports.py --xml caminho/para/cronograma.xml --out reports
```

Também é possível omitir `--xml` quando o arquivo XML estiver na mesma pasta do script e seguir o padrão `SRB_MIRAI_TP-0500_Cronograma*.xml`.
