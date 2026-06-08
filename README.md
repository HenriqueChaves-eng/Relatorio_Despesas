# Relatorio de Despesas Agres

Este repositorio guarda os dois componentes do processo:

- Coletor offline para iPhone e iPad pelo GitHub Pages.
- Gerador local de Excel e PDF para rodar no PC.

O coletor fica publicado em HTTPS para instalacao no iPhone/iPad. O gerador nao precisa ser
publicado na nuvem: ele roda no proprio computador.

## Enderecos

- Coletor offline: `https://henriquechaves-eng.github.io/Relatorio_Despesas/`
- Gerador local: `http://127.0.0.1:8502`

## Publicar o coletor offline

No GitHub, abra `Settings` > `Pages` e configure:

- Source: `Deploy from a branch`
- Branch: `main`
- Folder: `/(root)`

O GitHub Pages usa o arquivo `index.html` da raiz.

## Rodar o gerador no PC

1. Baixe ou atualize os arquivos deste repositorio no PC.
2. Execute `iniciar_app.bat`.
3. Abra `http://127.0.0.1:8502`.
4. Importe o ZIP gerado no iPhone/iPad.

O gerador usa Streamlit localmente. Nenhum comprovante precisa ser enviado para servidor externo.

## Fluxo de uso

1. Registre despesas e comprovantes no coletor instalado no iPhone/iPad.
2. Gere o ZIP do coletor.
3. Abra o gerador no PC e selecione o ZIP.
4. Revise os lancamentos.
5. Baixe o pacote final contendo o Excel preenchido e o PDF dos comprovantes.

## Privacidade

O coletor salva despesas e fotos localmente no aparelho. O ZIP e importado no proprio PC pelo
gerador local.
