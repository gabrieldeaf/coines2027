# COINES 2027: Inscrição e Administração

Arquivos:
- `Code.gs`: back-end no Google Apps Script (a planilha é o banco de dados)
- `index.html`: site do participante (CPF, cadastro, painel)
- `admin.html`: site da equipe (scanners e tabelas)

## 1. Planilha e Apps Script
1. Crie uma planilha no Google Sheets, por exemplo "COINES 2027 - Inscrições".
2. Na planilha: **Extensões > Apps Script**. Apague o conteúdo e cole o `Code.gs`.
3. Em **Configurações do projeto** (engrenagem), defina o fuso horário como `(GMT-03:00) São Paulo`.
4. Ajuste o bloco `CONFIG` no topo do `Code.gs`: dias do evento, texto do período, carga horária, assinatura, frequência mínima.
5. Selecione a função `setup` e clique em **Executar** (autorize na primeira vez). Ela cria as abas:
   `Inscritos`, `Presenca`, `Kit`, `Lanche`, `Minicursos`, `PresencaMinicurso`.
6. Senha da equipe: o `setup` cria `coines2027`. Troque em **Configurações do projeto > Propriedades do script > ADMIN_SENHA**.

## 2. Publicar a API
1. **Implantar > Nova implantação > App da Web**.
2. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**.
3. Copie a URL que termina em `/exec` e cole em `API_URL` no `index.html` e no `admin.html`.
4. Ao alterar o `Code.gs` depois: **Implantar > Gerenciar implantações > editar (lápis) > Versão: Nova versão**. Assim a URL continua a mesma.

## 3. Hospedar os sites
Use GitHub Pages (ou qualquer hospedagem com HTTPS, obrigatório para a câmera):
1. Crie um repositório e envie `index.html` e `admin.html`.
2. **Settings > Pages > Branch: main**.
3. Participante: `https://SEU-USUARIO.github.io/REPO/`
   Equipe: `https://SEU-USUARIO.github.io/REPO/admin.html`

## 4. Minicursos
- Preencha a aba `Minicursos` (ID, Título, Data `dd/mm/aaaa`, HoraInicio e HoraFim `hh:mm`, Local, CargaHoraria, Ministrante).
- No admin, aba **Tabelas > Minicursos**, botão **QR** em cada linha: imprima e cole na porta da sala.
- O participante escaneia esse QR pelo próprio painel. Com `VALIDAR_HORARIO_MINICURSO: true`, só vale no dia e no horário (com tolerância de 30 min). Para testar antes do evento, coloque `false`.

## Como funciona
| QR | Conteúdo | Onde é lido |
|---|---|---|
| Presencial | `COINES27\|P\|CODIGO` | Admin: Presença (1 por dia) e Kit (1 por pessoa) |
| Lanche | `COINES27\|L\|CODIGO` | Admin: Lanche (1 de manhã e 1 à tarde, por dia) |
| Minicurso | `COINES27\|M\|MC01` | Participante, no painel |

- **Frequência** = dias com presença ÷ dias do evento. Com 3 dias e mínimo de 75%, é preciso estar nos 3 dias. Com 67% bastam 2.
- **Certificado do evento** libera ao atingir a frequência mínima. **Certificado de minicurso** libera com a presença registrada. O PDF é gerado no navegador.
- Fundo personalizado do certificado: salve uma imagem A4 paisagem como `certificado-fundo.png` ao lado do `index.html` e preencha `CERT_FUNDO`.
- O QR usa um código aleatório, não o CPF.
- Se o QR não ler, a equipe pode digitar o CPF ou o código no admin.
