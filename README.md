# Melhor trabalho SAF 2026

Aplicacao Streamlit para operar a votacao de melhor apresentacao com o minimo de infra:

- seed lida de Excel ou CSV e reduzida para `ID`, `Título da sinopse`, `chave_autor`, `nome_autor`, `aprovado`, `DIA`
- formulario de votacao com autenticacao por CHAVE + senha (DATA_INICIO em `ddmmaaaa`)
- 4 perguntas de nota com 1 a 5 estrelas
- QR code e link individual por trabalho aprovado
- painel administrativo com ranking, votos brutos, auditoria e orientacoes de deploy
- rota publica de votacao por apresentacao usando URL indexada por ID

## Arquitetura recomendada

### Teste local

- Frontend: Streamlit
- Persistencia: SQLite local
- Seed: Excel original ou CSV gerado

### Producao em Streamlit Community Cloud

- Frontend: Streamlit Community Cloud
- Persistencia: Supabase
- Seed: arquivo `Best_work_award/seed_data.csv` versionado no repositorio

Motivo: o filesystem do Streamlit Community Cloud e efemero. SQLite funciona bem para testes locais, mas nao deve ser a base oficial do evento na nuvem.

## Como rodar localmente

```powershell
conda activate saf
pip install -r Best_work_award\requirements.txt
streamlit run Best_work_award\app.py
```

## Variaveis de ambiente

### Comuns

- `BEST_WORK_AWARD_BASE_URL`: URL usada para gerar os links e QR codes
- `BEST_WORK_AWARD_ADMIN_USERNAME`: usuario do painel admin (padrao `admin_saf`)
- `BEST_WORK_AWARD_ADMIN_PASSWORD`: senha do painel admin (padrao `14159265`)
- `BEST_WORK_AWARD_SEED_PATH`: caminho da seed local ou versionada no repo
- `BEST_WORK_AWARD_ATTENDEES_PATH`: caminho opcional da planilha de participantes (padrao `Best_work_award/attendees_saf26.xlsx`)

Voce pode definir essas variaveis em `Best_work_award/.env`.

### Backend local

- `BEST_WORK_AWARD_STORAGE_BACKEND=sqlite`

### Backend Streamlit Community Cloud

- `BEST_WORK_AWARD_STORAGE_BACKEND=supabase`
- `BEST_WORK_AWARD_SUPABASE_URL`
- `BEST_WORK_AWARD_SUPABASE_KEY`
- `BEST_WORK_AWARD_SUPABASE_VOTES_TABLE` opcional, padrao `votes`

## Supabase

Execute o script [Best_work_award/supabase_schema.sql](Best_work_award/supabase_schema.sql) no SQL Editor do Supabase antes do deploy.

## Deploy no Streamlit Community Cloud

1. Suba este repositorio para o GitHub.
2. Garanta que `Best_work_award/seed_data.csv` esteja commitado no repositorio.
3. Garanta que `Best_work_award/attendees_saf26.xlsx` esteja disponivel no deploy (ou configure `BEST_WORK_AWARD_ATTENDEES_PATH`).
4. Crie um projeto no Supabase e rode [Best_work_award/supabase_schema.sql](Best_work_award/supabase_schema.sql).
5. No Streamlit Community Cloud, crie um novo app apontando para `Best_work_award/app.py`.
6. Configure os secrets com base em [Best_work_award/.streamlit/secrets.example.toml](Best_work_award/.streamlit/secrets.example.toml).
7. Defina `BEST_WORK_AWARD_BASE_URL` com a URL final do app publicado.
8. Abra a aba `Admin`, gere os QR codes e distribua no evento.

## URL individual por apresentacao

Formato da rota de voto individual:

`https://seu-app.streamlit.app?view=vote&id=75`

- `id` e o ID da apresentacao na seed
- essa rota abre uma pagina limpa para o usuario, focada no formulario de voto
- os QR codes gerados no painel admin ja usam esse formato
- a raiz do app (`/`) fica protegida por login de administrador
- somente as rotas individuais `?view=vote&id=...` ficam abertas para votacao publica

## Regras atuais de validacao e auditoria

- cada CHAVE vota apenas uma vez por trabalho
- autoavaliacao bloqueada
- CHAVE deve seguir o regex exigido e existir na planilha de participantes
- senha deve ser a DATA_INICIO da pessoa no formato `ddmmaaaa`
- aba de auditoria mostra volume de votos por CHAVE

## Instrucoes para o participante

- informe sua CHAVE corporativa
- informe sua senha no formato `ddmmaaaa`, usando a DATA_INICIO do cadastro
- exemplo: DATA_INICIO `17/07/2006` vira senha `17072006`
- apenas CHAVEs presentes em `attendees_saf26.xlsx` podem votar