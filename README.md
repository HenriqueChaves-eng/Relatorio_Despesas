# Publicar o Coletor Agres no GitHub Pages

O coletor funciona como um aplicativo web instalavel no iPhone e iPad. Os lancamentos
e as fotos ficam armazenados localmente no aparelho e nao sao enviados ao GitHub.

## Criar o repositorio

1. Acesse `https://github.com/new`.
2. Use o nome `Relatorio_Despesas`.
3. Escolha `Public`.
4. Nao marque as opcoes para criar README, `.gitignore` ou licenca.
5. Clique em `Create repository`.

## Enviar os arquivos

1. Na pagina do repositorio, clique em `uploading an existing file`.
2. Abra a pasta `PUBLICAR_GITHUB_PAGES_COLETOR_AGRES`.
3. Arraste todos os arquivos da pasta para a area de upload do GitHub.
4. Clique em `Commit changes`.

## Ativar o GitHub Pages

1. No repositorio, abra `Settings`.
2. No menu lateral, abra `Pages`.
3. Em `Build and deployment`, escolha `Deploy from a branch`.
4. Selecione a branch `main`, a pasta `/(root)` e clique em `Save`.
5. Aguarde a publicacao.

O endereco publicado sera:

`https://henriquechaves-eng.github.io/Relatorio_Despesas/`

## Instalar no iPhone ou iPad

1. Abra o endereco publicado usando o Safari.
2. Toque em `Compartilhar`.
3. Toque em `Adicionar a Tela de Inicio`.
4. Ative `Abrir como App`, se a opcao aparecer.
5. Abra o aplicativo uma vez com internet.
6. Ative o modo aviao e abra novamente para validar o funcionamento offline.

## Cuidados importantes

- O codigo do coletor sera publico, mas os lancamentos e comprovantes permanecem no aparelho.
- Exporte o pacote ZIP regularmente. Apagar os dados do Safari ou remover o aplicativo pode
  apagar os lancamentos ainda nao exportados.
- Para atualizar o coletor, envie os novos arquivos ao mesmo repositorio. O aplicativo
  instalara a nova versao quando voltar a ficar online.
