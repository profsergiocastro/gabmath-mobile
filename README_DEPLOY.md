GabMath Mobile Static

Esta pasta pode ser publicada como site estatico em qualquer hospedagem continua:
- Netlify
- GitHub Pages
- Cloudflare Pages
- Vercel

Arquivos necessarios:
- `index.html`
- `app.js`
- `styles.css`
- `netlify.toml` (opcional, util no Netlify)

Como funciona:
- o QR da prova ja carrega os dados minimos da avaliacao
- a pagina web le o QR
- a pagina reconstrui localmente o gabarito
- a correcao funciona sem backend e sem depender do PC ligado

Limite atual:
- a leitura automatica das bolhas por foto ainda nao foi implementada
- nesta etapa, o QR e lido pela camera e as alternativas sao marcadas manualmente no celular

Publicacao no Netlify:
1. Entre no Netlify.
2. Crie um novo site por upload manual ou por repositorio Git.
3. Use esta pasta `mobile_web` como pasta publicada.
4. Depois do primeiro deploy, o endereco fica fixo.

Observacao:
- para um dominio realmente fixo, a publicacao precisa ficar vinculada a uma conta sua no provedor de hospedagem
- sem isso, eu consigo gerar apenas links temporarios
