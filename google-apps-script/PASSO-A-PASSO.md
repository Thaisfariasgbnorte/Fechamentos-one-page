# Colocar o Comparativo F&I no ar pelo Google (uma vez só, ~5 minutos)

Quem faz: você (Thais), com a conta Google da GB Norte. Depois disso, a sua chefe só usa o link.

## 1. Criar a planilha do histórico
1. Abra https://sheets.new e dê o nome **Histórico Comparativo F&I**.
2. Menu **Extensões → Apps Script**.

## 2. Colar o código
1. No Apps Script, apague o que estiver em `Código.gs` e cole todo o conteúdo do arquivo **Code.gs**.
2. Clique no **+** ao lado de "Arquivos" → **HTML**, dê o nome **Index** (sem ".html") e cole
   todo o conteúdo do arquivo **Index.html** (apague o que vier escrito antes).
3. Clique em **Salvar** (ícone de disquete).

## 3. Criar as abas e autorizar
1. No topo, escolha a função **configurar** e clique em **Executar**.
2. O Google pede autorização: **Revisar permissões** → escolha sua conta → **Permitir**.
3. Na planilha aparecem as abas **FECHAMENTOS** e **EDITORES**.
4. Na aba **EDITORES**, abaixo do seu e-mail, escreva o e-mail da sua chefe (um por linha).
   Só quem está nessa lista consegue salvar ou excluir fechamentos.

## 4. Publicar o site
1. **Implantar → Nova implantação** → engrenagem → **App da Web**.
2. Descrição: `Comparativo F&I`.
3. **Executar como:** Eu. **Quem pode acessar:** Qualquer pessoa em GB Norte (seu domínio).
4. **Implantar** e copie o **URL do app da Web** (termina em `/exec`). Esse é o link do site:
   mande para a sua chefe e para quem precisar ver.

## Uso no dia a dia (sua chefe)
1. Abrir o link (entrar com o e-mail da GB Norte se pedir).
2. **Baixar planilha base** → atualizar no Excel → **Enviar planilha atualizada**.
3. Conferir a prévia, dar um nome e clicar **Salvar como novo fechamento**.
4. **Gerar PDF** baixa o fechamento que está na tela.
5. Para ver um fechamento antigo, escolha no filtro **Fechamento**.

Cada fechamento é uma linha na aba FECHAMENTOS e nunca é alterado depois de salvo.
Não edite essa aba à mão (os dados ficam em JSON nas colunas G em diante).

## Se mudar o código depois
Cole a nova versão e use **Implantar → Gerenciar implantações → editar (lápis) → Versão: Nova versão
→ Implantar**. Assim o link continua o mesmo.
