from __future__ import annotations

import base64
import datetime as dt
import hashlib
import io
import os
import re
import sqlite3
import unicodedata
from pathlib import Path
from typing import Any, cast
from urllib.parse import urlencode
from zoneinfo import ZoneInfo

import pandas as pd
import qrcode
import requests
import streamlit as st
from dotenv import load_dotenv


APP_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = APP_DIR.parent
load_dotenv(APP_DIR / ".env")
LOCAL_EXCEL_PATH = (
	WORKSPACE_DIR / "SAF26" / "SAF2026_sinopses_resultado_final_forms_melhor_trabalho.xlsx"
)
DEFAULT_SEED_CSV_PATH = APP_DIR / "seed_data.csv"
DEFAULT_DB_PATH = APP_DIR / "best_work_award.db"
VOTE_HEADER_IMAGE_PATH = APP_DIR / "saf.png"
DEFAULT_ATTENDEES_PATH = APP_DIR / "attendees_saf26.xlsx"
DEFAULT_PHOTOS_PATH = APP_DIR / "photos"
RIO_TZ = ZoneInfo("America/Sao_Paulo")
SEED_COLUMNS = [
	"ID",
	"Título da sinopse",
	"chave_autor",
	"nome_autor",
	"aprovado",
	"DIA",
]
VOTE_COLUMNS = [
	"id",
	"work_id",
	"work_title",
	"author_key",
	"author_name",
	"presentation_day",
	"voter_key",
	"clarity_score",
	"engagement_score",
	"relevance_score",
	"overall_score",
	"average_score",
	"submitted_at",
]
VOTER_KEY_PATTERN = re.compile(r"^[A-Z][A-Z0-9]{3}$")
PIN_PATTERN = re.compile(r"^\d{8}$")
DEFAULT_ADMIN_USERNAME = "admin_saf"
DEFAULT_ADMIN_PASSWORD = "14159265"
ADMIN_AUTH_SESSION_KEY = "best_work_award_admin_ok"
PIN_CACHE_SESSION_KEY = "best_work_award_pin_cache"
VOTER_KEY_CACHE_SESSION_KEY = "best_work_award_last_voter_key"
VOTER_KEY_INPUT_WIDGET_KEY = "best_work_award_voter_key_input"
VOTER_AUTH_SESSION_KEY = "best_work_award_authenticated_voter_key"
DEFAULT_TECHNICAL_VOTER_KEYS = [
	"IM96",
	"CTUB",
	"FN14",
	"UR8Z",
	"CJ54",
	"CJ95",
	"BETM",
	"CTQ4",
	"CXL1",
	"CJ79",
]


def default_seed_path() -> Path:
	if DEFAULT_SEED_CSV_PATH.exists():
		return DEFAULT_SEED_CSV_PATH
	return LOCAL_EXCEL_PATH


def page_config(show_default_heading: bool = True) -> None:
	st.set_page_config(
		page_title="SAF 2026 | Melhor Trabalho",
		page_icon="🏆",
		layout="wide",
		initial_sidebar_state="expanded",
	)
	inject_styles()
	if show_default_heading:
		st.title("SAF 2026 | 🏆 Avaliação de Melhor Trabalho")
		st.caption(
			"Seed da programação, formulário de avaliação, QR codes e apuração em um único app."
		)


def inject_styles() -> None:
	st.markdown(
		"""
		<style>
		:root {
			--brand-ink: #14323f;
			--brand-accent: #df5b2d;
			--brand-sand: #f4efe5;
			--brand-sky: #dce8ec;
			--brand-mint: #d9efe7;
		}
		.stApp {
			background:
				radial-gradient(circle at top right, rgba(223, 91, 45, 0.18), transparent 32%),
				linear-gradient(180deg, #fbf8f1 0%, #f4efe5 100%);
		}
		.block-container {
			padding-top: 2rem;
			padding-bottom: 2rem;
		}
		div[data-testid="stMetric"] {
			background: rgba(255, 255, 255, 0.72);
			border: 1px solid rgba(20, 50, 63, 0.08);
			padding: 0.8rem 1rem;
			border-radius: 18px;
			box-shadow: 0 12px 26px rgba(20, 50, 63, 0.07);
		}
		.hero-card {
			padding: 1.4rem 1.5rem;
			border-radius: 24px;
			background: linear-gradient(135deg, rgba(20, 50, 63, 0.96), rgba(27, 87, 105, 0.92));
			color: #ffffff;
			box-shadow: 0 20px 40px rgba(20, 50, 63, 0.18);
			margin-bottom: 1rem;
		}
		.hero-card h2 {
			margin: 0 0 0.4rem 0;
			font-size: 1.8rem;
		}
		.hero-card p {
			margin: 0;
			font-size: 1rem;
			line-height: 1.45;
		}
		.pill-row {
			display: flex;
			gap: 0.5rem;
			flex-wrap: wrap;
			margin-top: 0.9rem;
		}
		.pill {
			padding: 0.35rem 0.8rem;
			border-radius: 999px;
			background: rgba(255, 255, 255, 0.12);
			border: 1px solid rgba(255, 255, 255, 0.2);
			font-size: 0.9rem;
		}
		.info-panel {
			padding: 1rem 1.1rem;
			border-radius: 20px;
			background: rgba(255, 255, 255, 0.72);
			border: 1px solid rgba(20, 50, 63, 0.08);
			box-shadow: 0 10px 24px rgba(20, 50, 63, 0.06);
			margin-bottom: 1rem;
		}
		.score-guide {
			padding: 0.85rem 1rem;
			border-radius: 18px;
			background: linear-gradient(135deg, rgba(217, 239, 231, 0.95), rgba(220, 232, 236, 0.95));
			border: 1px solid rgba(20, 50, 63, 0.08);
			margin-bottom: 1rem;
		}
		.kpi-card {
			padding: 1rem 1.1rem;
			border-radius: 20px;
			background: rgba(255, 255, 255, 0.78);
			border: 1px solid rgba(20, 50, 63, 0.08);
			margin-bottom: 1rem;
		}
		.kpi-card h4,
		.kpi-card p {
			margin: 0;
		}
		.kpi-card p:last-child {
			margin-top: 0.35rem;
			font-size: 1.35rem;
			font-weight: 700;
			color: var(--brand-ink);
		}
		</style>
		""",
		unsafe_allow_html=True,
	)


def render_hero(approved_count: int, backend_label: str) -> None:
	st.markdown(
		f"""
		<div class="hero-card">
			<h2>Votação enxuta para o evento, com apuração pronta para decisão.</h2>
			<p>
				Os participantes acessam um QR code por trabalho, informam a CHAVE e avaliam em 2 critérios de 1 a 5 estrelas.
				O painel administrativo consolida ranking, auditoria e exportação.
			</p>
			<div class="pill-row">
				<span class="pill">{approved_count} trabalhos aprovados</span>
				<span class="pill">Backend: {backend_label}</span>
				<span class="pill">Deploy local ou Streamlit Community Cloud</span>
			</div>
		</div>
		""",
		unsafe_allow_html=True,
	)


def image_to_data_uri(image_path: Path) -> str:
	encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
	return f"data:image/png;base64,{encoded}"


def render_vote_header() -> None:
	logo_markup = ""
	if VOTE_HEADER_IMAGE_PATH.exists():
		logo_markup = (
			f"<img src='{image_to_data_uri(VOTE_HEADER_IMAGE_PATH)}' "
			"style='width: 205px; max-width: 100%; height: auto; display: block; flex: 0 0 auto;' alt='SAF logo'/>"
		)

	st.markdown(
		f"""
		<style>
			.block-container {{
				padding-top: 0 !important;
				padding-bottom: 0.4rem !important;
			}}
			div[data-testid="stAppViewContainer"] > .main .block-container {{
				padding-top: 0 !important;
			}}
			div[data-testid="stVerticalBlock"] {{
				gap: 0 !important;
			}}
			div[data-testid="stMarkdownContainer"] p {{
				margin-bottom: 0 !important;
			}}
		</style>
		<div style="display: flex; align-items: center; gap: 0.85rem; margin: -0.8rem 0 -0.2rem 0; width: 100%;">
			{logo_markup}
			<h1 style="margin: 0; color: #14323f; font-size: 2.12rem; line-height: 0.98; font-weight: 700; white-space: nowrap;">
				SAF 2026 | 🏆 Avaliação de Melhor Trabalho
			</h1>
		</div>
		""",
		unsafe_allow_html=True,
	)
	st.markdown("<div style='height: 0;'></div>", unsafe_allow_html=True)


def storage_backend() -> str:
	return os.getenv("BEST_WORK_AWARD_STORAGE_BACKEND", "sqlite").strip().lower()


def storage_backend_label() -> str:
	backend = storage_backend()
	return "Supabase" if backend == "supabase" else "SQLite local"


def supabase_table_name() -> str:
	return os.getenv("BEST_WORK_AWARD_SUPABASE_VOTES_TABLE", "votes")


def supabase_rest_url() -> str:
	url = os.getenv("BEST_WORK_AWARD_SUPABASE_URL", "").strip()
	if not url:
		raise RuntimeError(
			"Defina BEST_WORK_AWARD_SUPABASE_URL e BEST_WORK_AWARD_SUPABASE_KEY para usar o backend Supabase."
		)
	return f"{url.rstrip('/')}/rest/v1/{supabase_table_name()}"


def supabase_headers() -> dict[str, str]:
	key = os.getenv("BEST_WORK_AWARD_SUPABASE_KEY", "").strip()
	if not key:
		raise RuntimeError(
			"Defina BEST_WORK_AWARD_SUPABASE_URL e BEST_WORK_AWARD_SUPABASE_KEY para usar o backend Supabase."
		)
	return {
		"apikey": key,
		"Authorization": f"Bearer {key}",
		"Content-Type": "application/json",
		"Prefer": "return=representation",
	}


def read_source_frame(seed_source: Any, sheet_name: str) -> pd.DataFrame:
	if hasattr(seed_source, "name"):
		name = str(seed_source.name).lower()
		if name.endswith(".csv"):
			return pd.read_csv(seed_source)
		return pd.read_excel(seed_source, sheet_name=sheet_name)

	path = Path(str(seed_source))
	if path.suffix.lower() == ".csv":
		return pd.read_csv(path)
	return pd.read_excel(path, sheet_name=sheet_name)


def prepare_seed_data(df: pd.DataFrame) -> pd.DataFrame:
	missing_columns = [column for column in SEED_COLUMNS if column not in df.columns]
	if missing_columns:
		raise ValueError(
			"A planilha não contém as colunas esperadas: " + ", ".join(missing_columns)
		)

	seed = df.loc[:, SEED_COLUMNS].copy()
	seed["ID"] = seed["ID"].astype(str).str.strip()
	seed["Título da sinopse"] = seed["Título da sinopse"].astype(str).str.strip()
	seed["chave_autor"] = seed["chave_autor"].astype(str).str.strip().str.upper()
	seed["nome_autor"] = seed["nome_autor"].astype(str).str.strip()
	seed["aprovado"] = seed["aprovado"].astype(str).str.strip().str.lower()
	seed["DIA"] = pd.to_numeric(seed["DIA"], errors="coerce")
	seed = seed.dropna(subset=["ID", "Título da sinopse"])
	seed = seed.drop_duplicates(subset=["ID"]).sort_values(["DIA", "ID"])
	return seed.reset_index(drop=True)


def load_seed_data(seed_source: Any, sheet_name: str = "Sheet1") -> pd.DataFrame:
	df = read_source_frame(seed_source, sheet_name)
	return prepare_seed_data(df)


def approved_seed(seed: pd.DataFrame) -> pd.DataFrame:
	approved = seed.loc[seed["aprovado"].eq("sim")].copy()
	approved["DIA"] = approved["DIA"].fillna("")
	return approved.reset_index(drop=True)


def init_sqlite(db_path: Path) -> None:
	with sqlite3.connect(db_path) as conn:
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS votes (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				work_id TEXT NOT NULL,
				work_title TEXT NOT NULL,
				author_key TEXT NOT NULL,
				author_name TEXT NOT NULL,
				presentation_day REAL,
				voter_key TEXT NOT NULL,
				clarity_score INTEGER NOT NULL,
				engagement_score INTEGER NOT NULL,
				relevance_score INTEGER NOT NULL,
				overall_score INTEGER NOT NULL,
				average_score REAL NOT NULL,
				submitted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				UNIQUE(voter_key, work_id)
			)
			"""
		)
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS voter_pin_validation (
				voter_key TEXT PRIMARY KEY,
				validated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				validated_until TEXT NOT NULL
			)
			"""
		)
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS voter_key_prefill (
				client_hash TEXT PRIMARY KEY,
				voter_key TEXT NOT NULL,
				updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
			)
			"""
		)
		conn.execute(
			"""
			CREATE TABLE IF NOT EXISTS admin_auth_validation (
				client_hash TEXT PRIMARY KEY,
				validated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				validated_until TEXT NOT NULL
			)
			"""
		)


def init_storage(db_path: Path) -> None:
	backend = storage_backend()
	if backend == "supabase":
		supabase_rest_url()
		supabase_headers()
		return
	init_sqlite(db_path)


def empty_votes_frame() -> pd.DataFrame:
	return pd.DataFrame(columns=VOTE_COLUMNS)


def load_votes(db_path: Path) -> pd.DataFrame:
	backend = storage_backend()
	if backend == "supabase":
		response = requests.get(
			supabase_rest_url(),
			headers=supabase_headers(),
			params={"select": "*", "order": "submitted_at.desc"},
			timeout=30,
		)
		response.raise_for_status()
		data = response.json() or []
		if not data:
			return empty_votes_frame()
		return pd.DataFrame(data)

	with sqlite3.connect(db_path) as conn:
		try:
			votes = pd.read_sql_query("SELECT * FROM votes", conn)
		except Exception:
			return empty_votes_frame()
	return votes if not votes.empty else empty_votes_frame()


def vote_payload(
	work: pd.Series,
	voter_key: str,
	clarity_score: int,
	engagement_score: int,
	relevance_score: int,
	overall_score: int,
) -> dict[str, Any]:
	average_score = round(
		(clarity_score + engagement_score + relevance_score + overall_score) / 4,
		4,
	)
	return {
		"work_id": str(work["ID"]),
		"work_title": str(work["Título da sinopse"]),
		"author_key": str(work["chave_autor"]),
		"author_name": str(work["nome_autor"]),
		"presentation_day": None if pd.isna(work["DIA"]) else float(work["DIA"]),
		"voter_key": voter_key,
		"clarity_score": clarity_score,
		"engagement_score": engagement_score,
		"relevance_score": relevance_score,
		"overall_score": overall_score,
		"average_score": average_score,
	}


def rio_now() -> dt.datetime:
	return dt.datetime.now(RIO_TZ)


def rio_timestamp_sql() -> str:
	return rio_now().strftime("%Y-%m-%d %H:%M:%S")


def rio_timestamp_iso() -> str:
	return rio_now().isoformat()


def format_vote_timestamp_for_display(value: Any) -> str:
	if value is None:
		return "data indisponivel"

	parsed = pd.to_datetime(value, errors="coerce")
	if pd.isna(parsed):
		return "data indisponivel"

	if getattr(parsed, "tzinfo", None) is None:
		parsed = parsed.tz_localize(RIO_TZ)
	else:
		parsed = parsed.tz_convert(RIO_TZ)

	return parsed.strftime("%d/%m/%Y %H:%M:%S")


def save_vote(
	db_path: Path,
	work: pd.Series,
	voter_key: str,
	clarity_score: int,
	engagement_score: int,
	relevance_score: int,
	overall_score: int,
) -> tuple[bool, str]:
	payload = vote_payload(
		work=work,
		voter_key=voter_key,
		clarity_score=clarity_score,
		engagement_score=engagement_score,
		relevance_score=relevance_score,
		overall_score=overall_score,
	)

	backend = storage_backend()
	if backend == "supabase":
		try:
			supabase_payload = payload.copy()
			supabase_payload["submitted_at"] = rio_timestamp_iso()
			headers = supabase_headers().copy()
			headers["Prefer"] = "resolution=merge-duplicates,return=representation"
			response = requests.post(
				supabase_rest_url(),
				headers=headers,
				params={"on_conflict": "voter_key,work_id"},
				json=supabase_payload,
				timeout=30,
			)
			response.raise_for_status()
			return True, "Voto registrado/atualizado com sucesso."
		except requests.HTTPError as exc:
			return False, "Falha ao registrar o voto no backend configurado."

	try:
		submitted_at_rio = rio_timestamp_sql()
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				INSERT INTO votes (
					work_id,
					work_title,
					author_key,
					author_name,
					presentation_day,
					voter_key,
					clarity_score,
					engagement_score,
					relevance_score,
					overall_score,
					average_score,
					submitted_at
				) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				ON CONFLICT(voter_key, work_id) DO UPDATE SET
					work_title = excluded.work_title,
					author_key = excluded.author_key,
					author_name = excluded.author_name,
					presentation_day = excluded.presentation_day,
					clarity_score = excluded.clarity_score,
					engagement_score = excluded.engagement_score,
					relevance_score = excluded.relevance_score,
					overall_score = excluded.overall_score,
					average_score = excluded.average_score,
					submitted_at = excluded.submitted_at
				""",
				(
					payload["work_id"],
					payload["work_title"],
					payload["author_key"],
					payload["author_name"],
					payload["presentation_day"],
					payload["voter_key"],
					payload["clarity_score"],
					payload["engagement_score"],
					payload["relevance_score"],
					payload["overall_score"],
					payload["average_score"],
					submitted_at_rio,
				),
			)
		return True, "Voto registrado/atualizado com sucesso."
	except sqlite3.IntegrityError:
		return False, "Não foi possível registrar/atualizar o voto."


def build_vote_url(base_url: str, work_id: str) -> str:
	params = {"view": "vote", "id": str(work_id)}
	return f"{base_url.rstrip('/')}?{urlencode(params)}"


def qr_code_bytes(url: str) -> bytes:
	qr = qrcode.QRCode(box_size=6, border=2)
	qr.add_data(url)
	qr.make(fit=True)
	image = qr.make_image(fill_color="black", back_color="white")
	buffer = io.BytesIO()
	image.save(buffer, kind="PNG")
	return buffer.getvalue()


def format_stars(value: int) -> str:
	return f"{'★' * value}{'☆' * (5 - value)}"


def vote_question(label: str, key: str) -> int | None:
	return st.radio(
		label,
		options=[1, 2, 3, 4, 5],
		horizontal=True,
		format_func=format_stars,
		index=None,
		key=key,
	)


def latest_vote_for_work(db_path: Path, work_id: str, voter_key: str) -> pd.Series | None:
	votes = load_votes(db_path)
	if votes.empty:
		return None

	work_id_text = str(work_id).strip()
	voter_key_text = str(voter_key).strip().upper()
	votes_lookup = votes.copy()
	votes_lookup["work_id"] = votes_lookup["work_id"].astype(str).str.strip()
	votes_lookup["voter_key"] = votes_lookup["voter_key"].astype(str).str.strip().str.upper()

	match = votes_lookup.loc[
		votes_lookup["work_id"].eq(work_id_text)
		& votes_lookup["voter_key"].eq(voter_key_text)
	].copy()
	if match.empty:
		return None

	if "submitted_at" in match.columns:
		match["_submitted_at"] = pd.to_datetime(match["submitted_at"], errors="coerce")
		match = match.sort_values("_submitted_at")

	return match.iloc[-1]


def parse_voter_keys(raw_keys: str) -> set[str]:
	if not raw_keys.strip():
		return set()
	normalized = {
		key.strip().upper()
		for key in re.split(r"[;,\s]+", raw_keys)
		if key.strip()
	}
	return {key for key in normalized if VOTER_KEY_PATTERN.fullmatch(key)}


def ranking_dataframe(
	votes: pd.DataFrame,
	min_votes: int,
	technical_voter_keys: set[str] | None = None,
	technical_weight: float = 0.75,
	other_weight: float = 0.25,
) -> pd.DataFrame:
	if votes.empty:
		return pd.DataFrame()
	work_group_columns = [
		"work_id",
		"work_title",
		"author_name",
		"author_key",
		"presentation_day",
	]
	technical_voter_keys = {key.strip().upper() for key in (technical_voter_keys or set()) if key.strip()}
	votes_with_group = votes.copy()
	votes_with_group["voter_key"] = votes_with_group["voter_key"].astype(str).str.strip().str.upper()
	votes_with_group["is_technical_group"] = votes_with_group["voter_key"].isin(technical_voter_keys)

	ranking = (
		votes_with_group.groupby(
			work_group_columns,
			dropna=False,
		)
		.agg(
			quantidade_votos=("voter_key", "count"),
			media_clareza=("clarity_score", "mean"),
			media_envolvimento=("engagement_score", "mean"),
			media_relevancia=("relevance_score", "mean"),
			media_geral=("overall_score", "mean"),
		)
		.reset_index()
	)

	technical_scores = (
		votes_with_group.loc[votes_with_group["is_technical_group"]]
		.groupby(work_group_columns, dropna=False)
		.agg(
			votos_grupo_tecnico=("voter_key", "count"),
			nota_grupo_tecnico=("average_score", "mean"),
		)
		.reset_index()
	)

	other_scores = (
		votes_with_group.loc[~votes_with_group["is_technical_group"]]
		.groupby(work_group_columns, dropna=False)
		.agg(
			votos_demais_avaliadores=("voter_key", "count"),
			nota_demais_avaliadores=("average_score", "mean"),
		)
		.reset_index()
	)

	ranking = ranking.merge(technical_scores, on=work_group_columns, how="left")
	ranking = ranking.merge(other_scores, on=work_group_columns, how="left")
	ranking["votos_grupo_tecnico"] = ranking["votos_grupo_tecnico"].fillna(0).astype(int)
	ranking["votos_demais_avaliadores"] = ranking["votos_demais_avaliadores"].fillna(0).astype(int)

	technical_weight = max(0.0, float(technical_weight))
	other_weight = max(0.0, float(other_weight))

	def weighted_score(row: pd.Series) -> float:
		weighted_total = 0.0
		applied_weight = 0.0

		if row["votos_grupo_tecnico"] > 0 and pd.notna(row["nota_grupo_tecnico"]):
			weighted_total += technical_weight * float(row["nota_grupo_tecnico"])
			applied_weight += technical_weight

		if row["votos_demais_avaliadores"] > 0 and pd.notna(row["nota_demais_avaliadores"]):
			weighted_total += other_weight * float(row["nota_demais_avaliadores"])
			applied_weight += other_weight

		if applied_weight <= 0:
			return float("nan")

		# If only one group has votes, normalize by applied weights to avoid unfair penalization.
		return weighted_total / applied_weight

	ranking["nota_final"] = ranking.apply(weighted_score, axis=1)
	ranking = ranking.loc[ranking["quantidade_votos"] >= min_votes].copy()
	ranking = ranking.sort_values(
		["nota_final", "media_geral", "quantidade_votos", "work_title"],
		ascending=[False, False, False, True],
	)
	return ranking.reset_index(drop=True)


def audit_dataframe(votes: pd.DataFrame) -> pd.DataFrame:
	if votes.empty:
		return pd.DataFrame()

	audit = (
		votes.groupby("voter_key", dropna=False)
		.agg(
			votos_emitidos=("work_id", "count"),
			trabalhos_distintos=("work_id", "nunique"),
			dias_distintos=("presentation_day", "nunique"),
			nota_media_dada=("average_score", "mean"),
			ultima_votacao=("submitted_at", "max"),
		)
		.reset_index()
		.sort_values(["votos_emitidos", "ultima_votacao"], ascending=[False, False])
	)
	return audit


def voter_stats_dataframe(votes: pd.DataFrame) -> pd.DataFrame:
	if votes.empty:
		return pd.DataFrame()

	votes_stats = votes.copy()
	votes_stats["voter_key"] = votes_stats["voter_key"].astype(str).str.strip().str.upper()
	votes_stats["work_id"] = votes_stats["work_id"].astype(str).str.strip()
	votes_stats["average_score"] = pd.to_numeric(votes_stats["average_score"], errors="coerce")
	if "submitted_at" in votes_stats.columns:
		votes_stats["_submitted_at"] = pd.to_datetime(votes_stats["submitted_at"], errors="coerce")
	else:
		votes_stats["_submitted_at"] = pd.NaT

	# Keep only the latest vote for each voter/work pair so each work counts once.
	votes_stats = votes_stats.sort_values("_submitted_at")
	votes_stats = votes_stats.drop_duplicates(subset=["voter_key", "work_id"], keep="last")

	stats = (
		votes_stats.groupby("voter_key", dropna=False)
		.agg(
			trabalhos_avaliados=("work_id", "nunique"),
			media_notas=("average_score", "mean"),
			desvio_padrao_notas=("average_score", "std"),
			nota_minima=("average_score", "min"),
			nota_maxima=("average_score", "max"),
			ultima_votacao=("submitted_at", "max"),
		)
		.reset_index()
	)

	stats["media_notas"] = stats["media_notas"].round(3)
	stats["desvio_padrao_notas"] = stats["desvio_padrao_notas"].round(3)
	stats["nota_minima"] = stats["nota_minima"].round(3)
	stats["nota_maxima"] = stats["nota_maxima"].round(3)
	stats = stats.sort_values(["trabalhos_avaliados", "media_notas", "voter_key"], ascending=[False, False, True])
	return stats.reset_index(drop=True)


def allowed_voter_keys() -> set[str]:
	raw = os.getenv("BEST_WORK_AWARD_ALLOWED_VOTER_KEYS", "")
	if not raw.strip():
		return set()
	return {key.strip().upper() for key in raw.split(",") if key.strip()}


def attendees_path() -> Path:
	configured = os.getenv("BEST_WORK_AWARD_ATTENDEES_PATH", "").strip()
	if configured:
		return Path(configured)
	return DEFAULT_ATTENDEES_PATH


def photos_path() -> Path:
	configured = os.getenv("BEST_WORK_AWARD_PHOTOS_PATH", "").strip()
	if configured:
		return Path(configured)
	return DEFAULT_PHOTOS_PATH


def file_cache_token(path: Path) -> float | None:
	if not path.exists():
		return None
	return path.stat().st_mtime


def find_author_photo(author_key: str) -> Path | None:
	if not author_key:
		return None

	key = author_key.strip().upper()
	base_path = photos_path()
	for extension in (".jpg", ".jpeg", ".png", ".webp"):
		candidate = base_path / f"{key}{extension}"
		if candidate.exists():
			return candidate
	return None


def normalize_column_name(name: str) -> str:
	text = unicodedata.normalize("NFKD", str(name))
	text = "".join(char for char in text if not unicodedata.combining(char))
	text = text.upper().strip().replace(" ", "_")
	return text


def compact_column_name(name: str) -> str:
	return re.sub(r"[^A-Z0-9]", "", normalize_column_name(name))


def resolve_column_name(
	normalized_columns: dict[str, str],
	*candidates: str,
) -> str | None:
	for candidate in candidates:
		column = normalized_columns.get(candidate)
		if column:
			return column

	compact_candidates = {re.sub(r"[^A-Z0-9]", "", candidate) for candidate in candidates}
	for normalized_name, original_name in normalized_columns.items():
		if compact_column_name(normalized_name) in compact_candidates:
			return original_name

	return None


def normalize_chave_value(value: Any) -> str | None:
	if pd.isna(value):
		return None

	text = re.sub(r"\s+", "", str(value).strip().upper())
	if not text:
		return None

	return text if VOTER_KEY_PATTERN.fullmatch(text) else None


def date_to_pin(value: Any) -> str | None:
	if pd.isna(value):
		return None

	parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
	if pd.isna(parsed) and isinstance(value, (int, float)):
		parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
	if pd.isna(parsed):
		return None

	return parsed.strftime("%d%m%Y")


def pin_cache_ttl_days() -> int:
	raw = os.getenv("BEST_WORK_AWARD_PIN_CACHE_DAYS", "5").strip()
	try:
		ttl = int(raw)
	except ValueError:
		ttl = 5
	return max(1, ttl)


def is_pin_cached(voter_key: str) -> bool:
	if not voter_key:
		return False

	cache = st.session_state.get(PIN_CACHE_SESSION_KEY, {})
	expires_at = cache.get(voter_key)
	if not isinstance(expires_at, (int, float)):
		return False

	now_ts = dt.datetime.now(dt.timezone.utc).timestamp()
	if now_ts >= expires_at:
		cache.pop(voter_key, None)
		st.session_state[PIN_CACHE_SESSION_KEY] = cache
		return False

	return True


def cache_valid_pin(voter_key: str) -> None:
	if not voter_key:
		return

	ttl_days = pin_cache_ttl_days()
	expires_at = dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=ttl_days)
	cache = st.session_state.get(PIN_CACHE_SESSION_KEY, {})
	cache[voter_key] = expires_at.timestamp()
	st.session_state[PIN_CACHE_SESSION_KEY] = cache


def sync_cached_voter_key_from_input() -> None:
	raw_value = str(st.session_state.get(VOTER_KEY_INPUT_WIDGET_KEY, ""))
	voter_key = raw_value.strip().upper()
	if voter_key:
		st.session_state[VOTER_KEY_CACHE_SESSION_KEY] = voter_key


def is_voter_pin_persisted_valid(db_path: Path, voter_key: str) -> bool:
	if not voter_key or storage_backend() != "sqlite":
		return False

	now_utc_iso = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
	try:
		with sqlite3.connect(db_path) as conn:
			row = conn.execute(
				"""
				SELECT validated_until
				FROM voter_pin_validation
				WHERE voter_key = ?
				""",
				(voter_key,),
			).fetchone()
	except Exception:
		return False

	if row is None:
		return False

	validated_until = str(row[0])
	return validated_until > now_utc_iso


def persist_valid_pin(db_path: Path, voter_key: str) -> None:
	if not voter_key or storage_backend() != "sqlite":
		return

	ttl_days = pin_cache_ttl_days()
	now_utc = dt.datetime.now(dt.timezone.utc)
	validated_until = (now_utc + dt.timedelta(days=ttl_days)).strftime("%Y-%m-%d %H:%M:%S")
	validated_at = now_utc.strftime("%Y-%m-%d %H:%M:%S")

	try:
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				INSERT INTO voter_pin_validation (voter_key, validated_at, validated_until)
				VALUES (?, ?, ?)
				ON CONFLICT(voter_key) DO UPDATE SET
					validated_at = excluded.validated_at,
					validated_until = excluded.validated_until
				""",
				(voter_key, validated_at, validated_until),
			)
	except Exception:
		# Keep the flow resilient even if persistence fails for any reason.
		return


def client_fingerprint() -> str | None:
	try:
		raw_headers = dict(getattr(st.context, "headers", {}))
	except Exception:
		raw_headers = {}

	headers = {str(key).lower(): str(value) for key, value in raw_headers.items()}

	if not headers:
		return None

	ip = str(
		headers.get("x-forwarded-for")
		or headers.get("x-real-ip")
		or headers.get("remote-addr")
		or ""
	).split(",")[0].strip()
	user_agent = str(headers.get("user-agent") or headers.get("sec-ch-ua") or "").strip()
	accept_language = str(headers.get("accept-language") or "").strip()
	host = str(headers.get("host") or "").strip()

	if not ip and not user_agent and not accept_language and not host:
		return None

	raw = f"{ip}|{user_agent}|{accept_language}|{host}".encode("utf-8")
	return hashlib.sha256(raw).hexdigest()


def load_prefilled_voter_key_for_client(db_path: Path) -> str | None:
	if storage_backend() != "sqlite":
		return None

	client_hash = client_fingerprint()
	if not client_hash:
		return None

	try:
		with sqlite3.connect(db_path) as conn:
			row = conn.execute(
				"""
				SELECT voter_key
				FROM voter_key_prefill
				WHERE client_hash = ?
				""",
				(client_hash,),
			).fetchone()
	except Exception:
		return None

	if row is None:
		return None

	voter_key = str(row[0]).strip().upper()
	if not voter_key or not VOTER_KEY_PATTERN.fullmatch(voter_key):
		return None

	return voter_key


def persist_prefilled_voter_key_for_client(db_path: Path, voter_key: str) -> None:
	if storage_backend() != "sqlite":
		return

	client_hash = client_fingerprint()
	if not client_hash:
		return

	voter_key = voter_key.strip().upper()
	if not voter_key:
		return

	now_utc = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
	try:
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				INSERT INTO voter_key_prefill (client_hash, voter_key, updated_at)
				VALUES (?, ?, ?)
				ON CONFLICT(client_hash) DO UPDATE SET
					voter_key = excluded.voter_key,
					updated_at = excluded.updated_at
				""",
				(client_hash, voter_key, now_utc),
			)
	except Exception:
		return


def clear_prefilled_voter_key_for_client(db_path: Path) -> None:
	if storage_backend() != "sqlite":
		return

	client_hash = client_fingerprint()
	if not client_hash:
		return

	try:
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				DELETE FROM voter_key_prefill
				WHERE client_hash = ?
				""",
				(client_hash,),
			)
	except Exception:
		return


def validate_auth_credentials(voter_key: str, voter_pin: str, attendee_pins: dict[str, str]) -> str | None:
	if not VOTER_KEY_PATTERN.fullmatch(voter_key):
		return "A CHAVE informada é inválida. Use o padrão: uma letra seguida de 3 caracteres alfanuméricos."

	if not PIN_PATTERN.fullmatch(voter_pin):
		return "Senha inválida. Use a data de ingresso na Petrobras no formato ddmmaaaa."

	expected_pin = attendee_pins.get(voter_key)
	if expected_pin is None:
		return "Esta CHAVE não está autorizada para votar neste evento."

	if voter_pin != expected_pin:
		return "Senha inválida para esta CHAVE. Use sua data de ingresso na Petrobras no formato ddmmaaaa."

	return None


def get_authenticated_voter_key(db_path: Path) -> str | None:
	voter_key = str(st.session_state.get(VOTER_AUTH_SESSION_KEY, "")).strip().upper()
	if voter_key and (is_pin_cached(voter_key) or is_voter_pin_persisted_valid(db_path, voter_key)):
		return voter_key

	voter_key = str(st.session_state.get(VOTER_KEY_CACHE_SESSION_KEY, "")).strip().upper()
	if voter_key and (is_pin_cached(voter_key) or is_voter_pin_persisted_valid(db_path, voter_key)):
		st.session_state[VOTER_AUTH_SESSION_KEY] = voter_key
		return voter_key

	voter_key = load_prefilled_voter_key_for_client(db_path) or ""
	voter_key = voter_key.strip().upper()
	if voter_key and (is_pin_cached(voter_key) or is_voter_pin_persisted_valid(db_path, voter_key)):
		st.session_state[VOTER_KEY_CACHE_SESSION_KEY] = voter_key
		st.session_state[VOTER_AUTH_SESSION_KEY] = voter_key
		return voter_key

	st.session_state.pop(VOTER_AUTH_SESSION_KEY, None)
	return None


def render_voter_auth_gate(db_path: Path, attendee_pins: dict[str, str], vote_only: bool) -> str | None:
	authenticated = get_authenticated_voter_key(db_path)
	if authenticated:
		is_admin_user = bool(st.session_state.get(ADMIN_AUTH_SESSION_KEY)) or is_admin_persisted_valid(db_path)
		status_col, action_col = st.columns([0.78, 0.22])
		with status_col:
			st.success(f"Votando como CHAVE {authenticated}")
		with action_col:
			st.markdown("<div style='height: 0.15rem;'></div>", unsafe_allow_html=True)
			change_key_clicked = st.button(
				"Trocar CHAVE",
				key="change-voter-key",
				use_container_width=True,
				disabled=not is_admin_user,
				help=None if is_admin_user else "Disponível apenas para administrador.",
			)

		if change_key_clicked:
			st.session_state.pop(VOTER_AUTH_SESSION_KEY, None)
			st.session_state.pop(VOTER_KEY_CACHE_SESSION_KEY, None)
			st.session_state.pop(VOTER_KEY_INPUT_WIDGET_KEY, None)
			clear_prefilled_voter_key_for_client(db_path)
			st.rerun()
		bottom_gap = "0.55rem" if vote_only else "0.45rem"
		st.markdown(f"<div style='height: {bottom_gap};'></div>", unsafe_allow_html=True)
		return authenticated

	if vote_only:
		st.info("Autentique-se uma vez para votar nesta e nas próximas apresentações.")
	else:
		st.markdown(
			"""
			<div class="info-panel">
				<strong>Autenticação do participante</strong><br/>
				Informe CHAVE e senha uma única vez. Depois, a sessão permanece validada por alguns dias.
			</div>
			""",
			unsafe_allow_html=True,
		)

	# Add a subtle offset so the auth fields do not stick to the block above.
	spacer_height = "2.5rem" if vote_only else "0.45rem"
	st.markdown(f"<div style='height: {spacer_height};'></div>", unsafe_allow_html=True)

	with st.form("voter-auth-form"):
		voter_key_input = st.text_input("CHAVE", max_chars=4, key=VOTER_KEY_INPUT_WIDGET_KEY)
		voter_pin_input = st.text_input(
			"Senha (Data de ingresso na Petrobras no formato ddmmaaaa)",
			max_chars=8,
			type="password",
		)
		submitted_auth = st.form_submit_button("Validar acesso", use_container_width=True)

	if not submitted_auth:
		return None

	voter_key = voter_key_input.strip().upper()
	voter_pin = re.sub(r"\D", "", voter_pin_input)
	validation_error = validate_auth_credentials(voter_key, voter_pin, attendee_pins)
	if validation_error:
		st.error(validation_error)
		return None

	st.session_state[VOTER_AUTH_SESSION_KEY] = voter_key
	st.session_state[VOTER_KEY_CACHE_SESSION_KEY] = voter_key
	cache_valid_pin(voter_key)
	persist_valid_pin(db_path, voter_key)
	persist_prefilled_voter_key_for_client(db_path, voter_key)
	st.rerun()
	return None


@st.cache_data(show_spinner=False)
def load_attendee_pins(path_str: str, _cache_token: float | None = None) -> dict[str, str]:
	path = Path(path_str)
	if not path.exists():
		raise FileNotFoundError(f"Arquivo de participantes não encontrado: {path}")

	df = pd.read_excel(path)
	if df.empty:
		raise ValueError("A planilha de participantes está vazia.")

	normalized_columns = {normalize_column_name(column): column for column in df.columns}
	chave_column = resolve_column_name(normalized_columns, "CHAVE")
	data_inicio_column = resolve_column_name(normalized_columns, "DATA_INICIO")
	if not chave_column or not data_inicio_column:
		raise ValueError(
			"A planilha de participantes precisa conter as colunas CHAVE e DATA_INICIO."
		)

	attendees = df.loc[:, [chave_column, data_inicio_column]].copy()
	attendees["CHAVE"] = attendees[chave_column].apply(normalize_chave_value)
	attendees["PIN"] = attendees[data_inicio_column].apply(date_to_pin)
	attendees = attendees.loc[
		attendees["CHAVE"].notna()
		& attendees["PIN"].notna()
	].copy()

	if attendees.empty:
		raise ValueError(
			"Nenhum participante válido foi encontrado (CHAVE no padrão e DATA_INICIO válida)."
		)

	duplicated = attendees.groupby("CHAVE")["PIN"].nunique()
	inconsistent = duplicated.loc[duplicated > 1]
	if not inconsistent.empty:
		invalid_keys = ", ".join(inconsistent.index.tolist()[:10])
		raise ValueError(
			"Foram encontradas CHAVEs com mais de um DATA_INICIO no arquivo de participantes: "
			f"{invalid_keys}"
		)

	unique_attendees = attendees.drop_duplicates(subset=["CHAVE"], keep="first")
	return dict(zip(unique_attendees["CHAVE"], unique_attendees["PIN"]))


@st.cache_data(show_spinner=False)
def load_attendee_lotacao(path_str: str, _cache_token: float | None = None) -> dict[str, str]:
	path = Path(path_str)
	if not path.exists():
		return {}

	try:
		df = pd.read_excel(path)
	except Exception:
		return {}

	if df.empty:
		return {}

	normalized_columns = {normalize_column_name(column): column for column in df.columns}
	chave_column = resolve_column_name(normalized_columns, "CHAVE")
	lotacao_column = resolve_column_name(normalized_columns, "SIGLA_LOTACAO", "LOTACAO")
	if not chave_column or not lotacao_column:
		return {}

	attendees = df.loc[:, [chave_column, lotacao_column]].copy()
	attendees["CHAVE"] = attendees[chave_column].apply(normalize_chave_value)
	attendees["LOTACAO"] = attendees[lotacao_column].fillna("").astype(str).str.strip()
	attendees = attendees.loc[
		attendees["CHAVE"].notna()
		& attendees["LOTACAO"].ne("")
	].copy()
	attendees = attendees.drop_duplicates(subset=["CHAVE"], keep="first")
	return dict(zip(attendees["CHAVE"], attendees["LOTACAO"]))


def validate_voter_key(
	voter_key: str,
	voter_pin: str,
	work: pd.Series,
	attendee_pins: dict[str, str],
	pin_required: bool = True,
) -> str | None:
	if not VOTER_KEY_PATTERN.fullmatch(voter_key):
		return "A CHAVE informada é inválida. Use o padrão: uma letra seguida de 3 caracteres alfanuméricos."

	if voter_key == str(work["chave_autor"]):
		return "Autoavaliação não é permitida."

	expected_pin = attendee_pins.get(voter_key)
	if expected_pin is None:
		return "Esta CHAVE não está autorizada para votar neste evento."

	if not pin_required:
		return None

	if not PIN_PATTERN.fullmatch(voter_pin):
		return "Senha inválida. Use a data de ingresso (DATA_INICIO) no formato ddmmaaaa."

	if voter_pin != expected_pin:
		return "Senha inválida para esta CHAVE. Use sua DATA_INICIO no formato ddmmaaaa."

	return None


def work_brief_description(work: pd.Series) -> str:
	day_value = "X"
	if pd.notna(work["DIA"]) and work["DIA"] != "":
		day_value = str(int(work["DIA"]))
	return f'ID {work["ID"]} - Autor: "{work["nome_autor"]}" apresentado no dia {day_value}'


def base_url_input() -> str:
	default_base_url = os.getenv("BEST_WORK_AWARD_BASE_URL", "http://localhost:8501/")
	return st.sidebar.text_input(
		"Base URL para gerar links e QR codes",
		value=default_base_url,
		help="Use a URL publica ou interna onde este app sera disponibilizado.",
	)


def admin_cache_ttl_days() -> int:
	raw = os.getenv("BEST_WORK_AWARD_ADMIN_CACHE_DAYS", "5").strip()
	try:
		ttl = int(raw)
	except ValueError:
		ttl = 5
	return max(1, ttl)


def is_admin_persisted_valid(db_path: Path) -> bool:
	if storage_backend() != "sqlite":
		return False

	client_hash = client_fingerprint()
	if not client_hash:
		return False

	now_utc_iso = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
	try:
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				CREATE TABLE IF NOT EXISTS admin_auth_validation (
					client_hash TEXT PRIMARY KEY,
					validated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
					validated_until TEXT NOT NULL
				)
				"""
			)
			row = conn.execute(
				"""
				SELECT validated_until
				FROM admin_auth_validation
				WHERE client_hash = ?
				""",
				(client_hash,),
			).fetchone()
	except Exception:
		return False

	if row is None:
		return False

	validated_until = str(row[0])
	return validated_until > now_utc_iso


def persist_admin_auth(db_path: Path) -> None:
	if storage_backend() != "sqlite":
		return

	client_hash = client_fingerprint()
	if not client_hash:
		return

	ttl_days = admin_cache_ttl_days()
	now_utc = dt.datetime.now(dt.timezone.utc)
	validated_until = (now_utc + dt.timedelta(days=ttl_days)).strftime("%Y-%m-%d %H:%M:%S")
	validated_at = now_utc.strftime("%Y-%m-%d %H:%M:%S")

	try:
		with sqlite3.connect(db_path) as conn:
			conn.execute(
				"""
				CREATE TABLE IF NOT EXISTS admin_auth_validation (
					client_hash TEXT PRIMARY KEY,
					validated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
					validated_until TEXT NOT NULL
				)
				"""
			)
			conn.execute(
				"""
				INSERT INTO admin_auth_validation (client_hash, validated_at, validated_until)
				VALUES (?, ?, ?)
				ON CONFLICT(client_hash) DO UPDATE SET
					validated_at = excluded.validated_at,
					validated_until = excluded.validated_until
				""",
				(client_hash, validated_at, validated_until),
			)
	except Exception:
		return


def admin_access(db_path: Path) -> bool:
	admin_username = os.getenv("BEST_WORK_AWARD_ADMIN_USERNAME", DEFAULT_ADMIN_USERNAME).strip()
	admin_password = os.getenv("BEST_WORK_AWARD_ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD).strip()

	if st.session_state.get(ADMIN_AUTH_SESSION_KEY):
		return True

	if is_admin_persisted_valid(db_path):
		st.session_state[ADMIN_AUTH_SESSION_KEY] = True
		return True

	with st.sidebar:
		st.subheader("Acesso administrativo")
		typed_username = st.text_input("Usuário")
		typed_password = st.text_input("Senha", type="password")
		if typed_username and typed_password and typed_username == admin_username and typed_password == admin_password:
			st.session_state[ADMIN_AUTH_SESSION_KEY] = True
			persist_admin_auth(db_path)
			st.rerun()

	st.info("Informe usuário e senha de administrador para acessar QR codes, votos e ranking.")
	return False


def root_access(db_path: Path) -> bool:
	admin_username = os.getenv("BEST_WORK_AWARD_ADMIN_USERNAME", DEFAULT_ADMIN_USERNAME).strip()
	admin_password = os.getenv("BEST_WORK_AWARD_ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD).strip()

	if st.session_state.get(ADMIN_AUTH_SESSION_KEY):
		return True

	if is_admin_persisted_valid(db_path):
		st.session_state[ADMIN_AUTH_SESSION_KEY] = True
		return True

	st.header("Acesso restrito")
	st.write("A página principal está protegida. Use as credenciais de administrador para continuar.")
	with st.form("root-login"):
		typed_username = st.text_input("Usuário")
		typed_password = st.text_input("Senha", type="password")
		submitted = st.form_submit_button("Entrar")

	if submitted and typed_username == admin_username and typed_password == admin_password:
		st.session_state[ADMIN_AUTH_SESSION_KEY] = True
		persist_admin_auth(db_path)
		st.rerun()

	if submitted:
		st.error("Usuário ou senha inválidos.")

	return False


def requested_work_id() -> str | None:
	# Keep backward compatibility with old work_id links while preferring id.
	return st.query_params.get("id") or st.query_params.get("work_id")


def exit_vote_only_view() -> None:
	try:
		st.query_params.clear()
	except Exception:
		# Fallback for Streamlit versions where clear may not be available.
		for param_key in ("view", "id", "work_id"):
			try:
				st.query_params.pop(param_key, None)
			except Exception:
				continue
	st.rerun()


def redirect_to_vote_exit_page() -> None:
	try:
		st.query_params.clear()
	except Exception:
		for param_key in ("view", "id", "work_id"):
			try:
				st.query_params.pop(param_key, None)
			except Exception:
				continue
	st.query_params["view"] = "closed"
	st.rerun()


def render_vote_exit_page() -> None:
	st.markdown(
		"""
		<style>
			[data-testid="stSidebar"] {
				display: none;
			}
		</style>
		<div style="
			margin: 5.5rem auto 0 auto;
			max-width: 720px;
			padding: 1.4rem 1.2rem;
			border-radius: 18px;
			background: rgba(255, 255, 255, 0.82);
			border: 1px solid rgba(20, 50, 63, 0.12);
			box-shadow: 0 12px 28px rgba(20, 50, 63, 0.08);
			text-align: center;
		">
			<h2 style="margin: 0 0 0.6rem 0; color: #14323f;">Votação encerrada nesta aba</h2>
			<p style="margin: 0; color: #2f4852; font-size: 1.02rem; line-height: 1.45;">
				Você já pode fechar esta página, se desejar.
			</p>
		</div>
		""",
		unsafe_allow_html=True,
	)


def selected_work(approved: pd.DataFrame, force_id: str | None = None) -> pd.Series | None:
	current_work_id = force_id or requested_work_id()
	if current_work_id:
		matches = approved.loc[approved["ID"].eq(str(current_work_id))]
		if not matches.empty:
			return matches.iloc[0]
		if force_id is not None:
			return None

	filters_col1, filters_col2 = st.columns([1, 2])
	day_options = ["Todos"] + [
		str(int(day))
		for day in sorted(day for day in approved["DIA"].dropna().unique())
	]
	selected_day = filters_col1.selectbox("Filtrar por dia", day_options)
	search_text = filters_col2.text_input("Buscar por ID, autor ou título")

	filtered = approved.copy()
	if selected_day != "Todos":
		filtered = filtered.loc[filtered["DIA"].astype(str).eq(selected_day)]

	if search_text.strip():
		needle = search_text.strip().lower()
		filtered = filtered.loc[
			filtered["ID"].str.lower().str.contains(needle)
			| filtered["Título da sinopse"].str.lower().str.contains(needle)
			| filtered["nome_autor"].str.lower().str.contains(needle)
		]

	if filtered.empty:
		st.warning("Nenhum trabalho corresponde aos filtros atuais.")
		return None

	work_options = {
		f"ID {row['ID']} | Dia {int(row['DIA']) if pd.notna(row['DIA']) and row['DIA'] != '' else '-'} | {row['Título da sinopse']}": row["ID"]
		for _, row in filtered.iterrows()
	}
	selected_label = st.selectbox("Selecione a apresentação", list(work_options), label_visibility="collapsed")
	selected_id = work_options[selected_label]
	matches = filtered.loc[filtered["ID"].eq(str(selected_id))]
	if matches.empty:
		return None
	return matches.iloc[0]


def render_vote_page(approved: pd.DataFrame, db_path: Path, vote_only: bool = False) -> None:
	if vote_only:
		st.markdown(
			"""
			<style>
				div[data-testid="stForm"] {
					margin-top: 0 !important;
					padding-top: 0 !important;
				}
				div[data-testid="stTextInputRootElement"] {
					margin-top: 0 !important;
				}
				div.row-widget.stRadio > div {
					gap: 0.1rem !important;
				}
				div[data-testid="stRadio"] {
					margin-bottom: 0.15rem !important;
				}
				label[data-testid="stWidgetLabel"] {
					margin-bottom: 0.1rem !important;
				}
			</style>
			""",
			unsafe_allow_html=True,
		)
		render_vote_header()

	# Keep stars visibly larger in both vote-only and standard voting pages.
	st.markdown(
		"""
		<style>
			div[data-testid="stRadio"] label[data-testid="stWidgetLabel"] p,
			div[data-testid="stRadio"] label[data-testid="stWidgetLabel"] span {
				font-size: 1.18rem !important;
				line-height: 1.32 !important;
				font-weight: 600 !important;
			}
			div[data-testid="stRadio"] label[data-baseweb="radio"] div,
			div[data-testid="stRadio"] label[data-baseweb="radio"] span,
			div[data-testid="stRadio"] label[data-baseweb="radio"] p {
				font-size: 1.55rem !important;
				line-height: 1.1 !important;
			}
		</style>
		""",
		unsafe_allow_html=True,
	)

	work = selected_work(approved, force_id=requested_work_id() if vote_only else None)
	if work is None:
		if vote_only:
			st.error("Trabalho não existente")
		else:
			st.markdown(
				"""
				<div class="info-panel">
					<strong>Como funciona</strong><br/>
					1. Informe sua CHAVE.<br/>
					2. Informe sua senha: data de ingresso na Petrobras no formato ddmmaaaa (ex.: 17072006).<br/>
					3. Apenas pessoas inscritas no SAF 2026 podem votar.<br/>
					4. O voto mais recente de cada CHAVE por trabalho é o que vale. Autoavaliação é bloqueada.
				</div>
				""",
				unsafe_allow_html=True,
			)
			st.markdown(
				"""
				<div class="score-guide">
					<strong>Escala sugerida</strong><br/>
					1 = fraco, 3 = adequado, 5 = excelente.
				</div>
				""",
				unsafe_allow_html=True,
			)
		return

	try:
		attendees_file = attendees_path()
		attendee_pins = load_attendee_pins(
			str(attendees_file),
			file_cache_token(attendees_file),
		)
	except Exception as exc:
		st.error(f"Falha na validação de participantes: {exc}")
		return

	if not vote_only:
		st.markdown(
			"""
			<div class="info-panel">
				<strong>Como funciona</strong><br/>
				1. Informe sua CHAVE.<br/>
				2. Informe sua senha: data de ingresso na Petrobras no formato ddmmaaaa (ex.: 17072006).<br/>
				3. Apenas pessoas inscritas no SAF 2026 podem votar.<br/>
				4. O voto mais recente de cada CHAVE por trabalho é o que vale. Autoavaliação é bloqueada.
			</div>
			""",
			unsafe_allow_html=True,
		)
		st.markdown(
			"""
			<div class="score-guide">
				<strong>Escala sugerida</strong><br/>
				1 = fraco, 3 = adequado, 5 = excelente.
			</div>
			""",
			unsafe_allow_html=True,
		)

	if vote_only:
		st.markdown(
			f"""
			<h2 style="margin: -5.5rem 0 1.0rem 0; color: #14323f; font-size: 1.42rem; line-height: 1.03;">
				{work['Título da sinopse']}
			</h2>
			<p style="margin: -1.55rem 0 0.55rem 0; color: #4a6270; font-size: 0.95rem; line-height: 1.15;">
				{work_brief_description(work)}
			</p>
			""",
			unsafe_allow_html=True,
		)
	else:
		st.subheader(str(work["Título da sinopse"]))
	if not vote_only:
		st.caption(work_brief_description(work))
	else:
		st.markdown("<div style='height: 2.5rem;'></div>", unsafe_allow_html=True)
		if not get_authenticated_voter_key(db_path):
			st.info(
				"Informe sua CHAVE e use como senha sua data de ingresso na Petrobras no formato ddmmaaaa. "
				"Somente pessoas inscritas no SAF 2026 podem votar."
			)

	# Keep a modest offset so auth and question blocks feel connected.
	st.markdown("<div style='height: 1.6rem;'></div>", unsafe_allow_html=True)
	authenticated_voter_key = render_voter_auth_gate(db_path, attendee_pins, vote_only)
	if not authenticated_voter_key:
		return

	existing_vote = latest_vote_for_work(db_path, str(work["ID"]), authenticated_voter_key)
	if existing_vote is not None:
		st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)
		st.warning(
			"Você já avaliou este trabalho. Se enviar novamente, sua avaliação anterior será atualizada."
		)
		try:
			st.markdown("<div style='height: 1.45rem;'></div>", unsafe_allow_html=True)
			clarity_last = int(existing_vote.get("clarity_score", 0))
			relevance_last = int(existing_vote.get("relevance_score", 0))
			when_last = format_vote_timestamp_for_display(existing_vote.get("submitted_at"))
			st.caption(
				"Última avaliação: "
				f"clareza {clarity_last}, relevância {relevance_last}. "
				f"Registrada em: {when_last}."
			)
		except Exception:
			pass

		question_box_gap = "1.5rem" if existing_vote is not None else "1.0rem"
		st.markdown(f"<div style='height: {question_box_gap};'></div>", unsafe_allow_html=True)

	with st.form(f"vote-form-{work['ID']}", clear_on_submit=True):
		clarity_score = vote_question(
			"1 - Quão clara e bem organizada foi a apresentação?",
			key=f"clarity-{work['ID']}",
		)
		relevance_score = vote_question(
			"2 - Quão relevante ou impactante você considera o trabalho apresentado?",
			key=f"relevance-{work['ID']}",
		)
		if vote_only:
			submit_col, exit_col = st.columns([0.5, 0.2])
			with submit_col:
				submitted = st.form_submit_button("Enviar avaliação", use_container_width=True)
			with exit_col:
				exit_requested = st.form_submit_button("Sair", use_container_width=True)
		else:
			submitted = st.form_submit_button("Enviar avaliação", use_container_width=True)
			exit_requested = False

	if exit_requested:
		redirect_to_vote_exit_page()
		return

	if not submitted:
		return
	if any(score is None for score in (clarity_score, relevance_score)):
		st.error("Selecione uma nota de 1 a 5 para as duas perguntas antes de enviar a avaliação.")
		return
	clarity_score_value = cast(int, clarity_score)
	relevance_score_value = cast(int, relevance_score)
	# Keep schema compatibility while using only two questions.
	engagement_score_value = clarity_score_value
	overall_score_value = relevance_score_value
	voter_key = authenticated_voter_key
	autoeval_error = validate_voter_key(voter_key, "", work, attendee_pins, pin_required=False)
	if autoeval_error:
		st.error(autoeval_error)
		return

	success, message = save_vote(
		db_path=db_path,
		work=work,
		voter_key=voter_key,
		clarity_score=clarity_score_value,
		engagement_score=engagement_score_value,
		relevance_score=relevance_score_value,
		overall_score=overall_score_value,
	)
	if success:
		st.success(message)
		st.balloons()
		if not vote_only:
			st.info(
				f"Resumo enviado: clareza {clarity_score_value}, relevância {relevance_score_value}."
			)
	else:
		st.error(message)


def render_seed_tab(seed: pd.DataFrame, approved: pd.DataFrame, seed_source_label: str) -> None:
	st.subheader("Seed carregada")
	col1, col2, col3, col4 = st.columns(4)
	col1.metric("Total na seed", len(seed))
	col2.metric("Aprovados", len(approved))
	col3.metric("Dias mapeados", int(seed["DIA"].dropna().nunique()))
	col4.metric("Origem", seed_source_label)

	st.dataframe(seed, use_container_width=True, hide_index=True)
	st.download_button(
		"Baixar seed filtrada em CSV",
		seed.to_csv(index=False).encode("utf-8-sig"),
		file_name="best_work_award_seed.csv",
		mime="text/csv",
	)


def render_qr_tab(approved: pd.DataFrame, base_url: str) -> None:
	st.subheader("Links e QR codes")
	url_table = approved.copy()
	url_table["url_votacao"] = url_table["ID"].apply(
		lambda work_id: build_vote_url(base_url, work_id)
	)
	st.dataframe(url_table, use_container_width=True, hide_index=True)

	for _, work in approved.iterrows():
		vote_url = build_vote_url(base_url, work["ID"])
		with st.expander(f"ID {work['ID']} | {work['Título da sinopse']}"):
			st.write(f"Autor: {work['nome_autor']} ({work['chave_autor']})")
			st.write(f"URL: {vote_url}")
			qr_bytes = qr_code_bytes(vote_url)
			st.image(qr_bytes, width=180)
			st.download_button(
				label=f"Baixar QR ID {work['ID']}",
				data=qr_bytes,
				file_name=f"qr_best_work_{work['ID']}.png",
				mime="image/png",
				key=f"download-qr-{work['ID']}",
			)


def render_results_tab(db_path: Path) -> None:
	st.subheader("Apuração")
	votes = load_votes(db_path)
	if votes.empty:
		st.info("Nenhum voto registrado ainda.")
		return

	default_technical_keys = os.getenv(
		"BEST_WORK_AWARD_TECHNICAL_VOTER_KEYS",
		",".join(DEFAULT_TECHNICAL_VOTER_KEYS),
	)
	st.markdown("### Configuração de pesos por grupo")
	technical_keys_input = st.text_area(
		"CHAVEs do grupo técnico",
		value=default_technical_keys,
		help="Separe por vírgula, ponto e vírgula, espaço ou quebra de linha.",
	)
	weight_col1, weight_col2 = st.columns(2)
	technical_weight_pct = weight_col1.number_input(
		"Peso do grupo tecnico (%)",
		min_value=0.0,
		max_value=100.0,
		value=75.0,
		step=1.0,
	)
	other_weight_pct = weight_col2.number_input(
		"Peso dos demais avaliadores (%)",
		min_value=0.0,
		max_value=100.0,
		value=25.0,
		step=1.0,
	)
	technical_keys = parse_voter_keys(technical_keys_input)
	st.caption(
		f"{len(technical_keys)} CHAVEs válidas no grupo técnico. "
		"A nota final usa os pesos informados e normaliza automaticamente quando apenas um grupo tiver votos."
	)

	min_votes = st.number_input(
		"Número mínimo de votos para entrar no ranking",
		min_value=1,
		value=1,
		step=1,
	)
	ranking = ranking_dataframe(
		votes=votes,
		min_votes=int(min_votes),
		technical_voter_keys=technical_keys,
		technical_weight=technical_weight_pct / 100.0,
		other_weight=other_weight_pct / 100.0,
	)
	top3 = ranking.head(3)

	col1, col2, col3 = st.columns(3)
	col1.metric("Votos registrados", len(votes))
	col2.metric("Apresentações avaliadas", votes["work_id"].nunique())
	col3.metric("Avaliadores únicos", votes["voter_key"].nunique())

	st.markdown("### Top 3")
	if top3.empty:
		st.warning("Ainda não há apresentações suficientes para o ranking com o filtro atual.")
	else:
		for position, (_, row) in enumerate(top3.iterrows(), start=1):
			st.markdown(
				f"""
				<div class="kpi-card">
					<h4>{position}o lugar | ID {row['work_id']}</h4>
					<p>{row['work_title']}</p>
					<p>Nota final {row['nota_final']:.2f} com {int(row['quantidade_votos'])} votos</p>
				</div>
				""",
				unsafe_allow_html=True,
			)

	st.markdown("### Ranking completo")
	st.dataframe(ranking, use_container_width=True, hide_index=True)
	st.download_button(
		"Baixar ranking em CSV",
		ranking.to_csv(index=False).encode("utf-8-sig"),
		file_name="best_work_award_ranking.csv",
		mime="text/csv",
	)

	st.markdown("### Votos brutos")
	st.dataframe(
		votes.sort_values("submitted_at", ascending=False),
		use_container_width=True,
		hide_index=True,
	)
	st.download_button(
		"Baixar votos em CSV",
		votes.to_csv(index=False).encode("utf-8-sig"),
		file_name="best_work_award_votes.csv",
		mime="text/csv",
	)


def render_winner_spotlight(
	row: pd.Series,
	place_label: str,
	medal: str,
	lotacao_map: dict[str, str],
	photo_width: int = 260,
) -> None:
	author_key = str(row["author_key"])
	lotacao = lotacao_map.get(author_key, "Lotação não informada")
	photo_file = find_author_photo(author_key)

	headline_col, photo_col = st.columns([1.25, 1])
	with headline_col:
		st.markdown(
			f"""
			<div class="kpi-card">
				<h4 style="font-size: 1.32rem; margin-bottom: 0.25rem;">{medal} {place_label}</h4>
				<p style="font-size: 1.18rem; font-weight: 700; margin-top: 0.15rem;">{row['author_name']}</p>
				<p style="font-size: 1.0rem; margin-top: 0.35rem;"><strong>Trabalho:</strong> {row['work_title']}</p>
				<p style="font-size: 0.98rem; margin-top: 0.25rem;"><strong>CHAVE:</strong> {author_key}</p>
				<p style="font-size: 0.98rem; margin-top: 0.12rem;"><strong>Lotação:</strong> {lotacao}</p>
				<p style="font-size: 0.98rem; margin-top: 0.12rem;"><strong>Nota final:</strong> {row['nota_final']:.2f}</p>
			</div>
			""",
			unsafe_allow_html=True,
		)
	with photo_col:
		if photo_file is not None:
			st.image(str(photo_file), width=photo_width)
		else:
			st.warning(f"Foto não encontrada para CHAVE {author_key} em {photos_path()}.")


def render_podium_overview(top3: pd.DataFrame, lotacao_map: dict[str, str]) -> None:
	st.markdown("### Pódio final")
	st.markdown(
		"""
		<style>
		.podium-box {
			border-radius: 18px;
			padding: 0.9rem 0.8rem;
			background: rgba(255, 255, 255, 0.82);
			border: 1px solid rgba(20, 50, 63, 0.1);
			box-shadow: 0 10px 24px rgba(20, 50, 63, 0.08);
			text-align: center;
			min-height: 188px;
		}
		.podium-medal {
			font-size: 2.3rem;
			line-height: 1;
			margin-bottom: 0.35rem;
		}
		.podium-name {
			font-size: 1.02rem;
			font-weight: 700;
			line-height: 1.25;
		}
		.podium-work {
			font-size: 0.9rem;
			line-height: 1.25;
			margin-top: 0.35rem;
		}
		</style>
		""",
		unsafe_allow_html=True,
	)

	second_row = top3.iloc[1]
	first_row = top3.iloc[0]
	third_row = top3.iloc[2]
	col2, col1, col3 = st.columns([1, 1.15, 1])

	def podium_box(col: Any, row: pd.Series, medal: str, place: str) -> None:
		author_key = str(row["author_key"])
		lotacao = lotacao_map.get(author_key, "Lotação não informada")
		photo_file = find_author_photo(author_key)
		with col:
			if photo_file is not None:
				st.image(str(photo_file), width=170)
			st.markdown(
				f"""
				<div class="podium-box">
					<div class="podium-medal">{medal}</div>
					<div class="podium-name">{place} | {row['author_name']}</div>
					<div class="podium-work">{row['work_title']}</div>
					<div class="podium-work"><strong>Lotação:</strong> {lotacao}</div>
					<div class="podium-work"><strong>Nota:</strong> {row['nota_final']:.2f}</div>
				</div>
				""",
				unsafe_allow_html=True,
			)

	podium_box(col2, second_row, "🥈", "2o lugar")
	podium_box(col1, first_row, "🥇", "1o lugar")
	podium_box(col3, third_row, "🥉", "3o lugar")


def render_podium_page(db_path: Path) -> None:
	st.header("Pódio")
	if not admin_access(db_path):
		return

	votes = load_votes(db_path)
	if votes.empty:
		st.info("Nenhum voto registrado ainda.")
		return

	technical_keys = parse_voter_keys(
		os.getenv("BEST_WORK_AWARD_TECHNICAL_VOTER_KEYS", ",".join(DEFAULT_TECHNICAL_VOTER_KEYS))
	)
	ranking = ranking_dataframe(
		votes=votes,
		min_votes=1,
		technical_voter_keys=technical_keys,
		technical_weight=0.75,
		other_weight=0.25,
	)
	if len(ranking) < 3:
		st.warning("São necessários pelo menos 3 trabalhos no ranking para montar o pódio.")
		st.dataframe(ranking, use_container_width=True, hide_index=True)
		return

	attendees_file = attendees_path()
	lotacao_map = load_attendee_lotacao(
		str(attendees_file),
		file_cache_token(attendees_file),
	)
	top3 = ranking.head(3).reset_index(drop=True)

	step_key = "best_work_award_podium_step"
	if step_key not in st.session_state:
		st.session_state[step_key] = 0

	control_col1, control_col2 = st.columns([1, 1])
	with control_col1:
		if st.button("Iniciar/voltar apresentação", use_container_width=True):
			st.session_state[step_key] = 1
			st.rerun()
	with control_col2:
		if st.button("Reiniciar pódio", use_container_width=True):
			st.session_state[step_key] = 0
			st.rerun()

	step = int(st.session_state.get(step_key, 0))
	if step == 0:
		st.info("Clique em Iniciar/voltar apresentação para revelar o 3o lugar.")
		return

	st.markdown("### 3o lugar")
	render_winner_spotlight(top3.iloc[2], "3o lugar", "🥉", lotacao_map)
	if step == 1 and st.button("Ir para o 2o lugar", use_container_width=True):
		st.session_state[step_key] = 2
		st.rerun()
	if step < 2:
		return

	st.markdown("### 2o lugar")
	render_winner_spotlight(top3.iloc[1], "2o lugar", "🥈", lotacao_map)
	if step == 2 and st.button("Ir para o grande vencedor", use_container_width=True):
		st.session_state[step_key] = 3
		st.rerun()
	if step < 3:
		return

	st.markdown("### Grande vencedor")
	render_winner_spotlight(top3.iloc[0], "1o lugar", "🥇", lotacao_map, photo_width=300)
	if step == 3 and st.button("Mostrar pódio completo", use_container_width=True):
		st.session_state[step_key] = 4
		st.rerun()
	if step < 4:
		return

	render_podium_overview(top3, lotacao_map)


def render_audit_tab(db_path: Path) -> None:
	st.subheader("Auditoria")
	votes = load_votes(db_path)
	if votes.empty:
		st.info("Sem dados para auditoria.")
		return

	audit = audit_dataframe(votes)
	st.markdown(
		"""
		<div class="info-panel">
			<strong>Controles atuais</strong><br/>
			- CHAVE validada por regex e presença em attendees_saf26.xlsx.<br/>
			- Senha validada com DATA_INICIO no formato ddmmaaaa.<br/>
			- Uma CHAVE mantém apenas o voto mais recente em cada trabalho.<br/>
			- Autoavaliação bloqueada.<br/>
			- Lista autorizada vem do arquivo de participantes.
		</div>
		""",
		unsafe_allow_html=True,
	)
	st.dataframe(audit, use_container_width=True, hide_index=True)


def render_voter_stats_tab(db_path: Path) -> None:
	st.subheader("Estatísticas dos avaliadores")
	votes = load_votes(db_path)
	if votes.empty:
		st.info("Nenhum voto registrado ainda.")
		return

	stats = voter_stats_dataframe(votes)
	if stats.empty:
		st.info("Não foi possível consolidar as estatísticas dos avaliadores.")
		return

	st.markdown(
		"""
		<div class="info-panel">
			<strong>Como as métricas são calculadas</strong><br/>
			- Cada trabalho conta no máximo uma vez por CHAVE (considera o voto mais recente).<br/>
			- Média e desvio padrão usam a nota média de cada avaliação (average_score).<br/>
			- Desvio padrão mais alto indica avaliador mais variável nas notas.
		</div>
		""",
		unsafe_allow_html=True,
	)

	col1, col2, col3 = st.columns(3)
	col1.metric("Avaliadores únicos", int(stats["voter_key"].nunique()))
	col2.metric("Média de trabalhos por avaliador", f"{stats['trabalhos_avaliados'].mean():.2f}")
	col3.metric("Desvio padrão médio", f"{stats['desvio_padrao_notas'].fillna(0).mean():.3f}")

	st.dataframe(stats, use_container_width=True, hide_index=True)
	st.download_button(
		"Baixar estatísticas dos avaliadores em CSV",
		stats.to_csv(index=False).encode("utf-8-sig"),
		file_name="best_work_award_voter_stats.csv",
		mime="text/csv",
	)


def render_deploy_tab(seed_source_label: str) -> None:
	st.subheader("Deploy")
	backend = storage_backend()
	seed_hint = "CSV no repositório" if seed_source_label.endswith(".csv") else seed_source_label
	st.markdown(
		f"""
		<div class="info-panel">
			<strong>Configuração atual</strong><br/>
			Backend ativo: {storage_backend_label()}<br/>
			Origem da seed: {seed_hint}<br/>
			Para Streamlit Community Cloud, a combinação recomendada é <strong>seed CSV versionada + Supabase</strong>.
		</div>
		""",
		unsafe_allow_html=True,
	)
	if backend == "sqlite":
		st.warning(
			"SQLite é ideal para teste local. Em Streamlit Community Cloud, ele não é adequado como base definitiva porque o filesystem é efêmero."
		)
	else:
		st.success("Backend pronto para deploy em Streamlit Community Cloud.")

	st.code(
		"""BEST_WORK_AWARD_STORAGE_BACKEND=supabase
BEST_WORK_AWARD_SUPABASE_URL=https://SEU-PROJETO.supabase.co
BEST_WORK_AWARD_SUPABASE_KEY=cole-a-service-role-ou-chave-segura
BEST_WORK_AWARD_ADMIN_PASSWORD=defina-uma-senha
BEST_WORK_AWARD_BASE_URL=https://seu-app.streamlit.app
BEST_WORK_AWARD_SEED_PATH=Best_work_award/seed_data.csv""",
		language="bash",
	)


def render_admin_page(
	seed: pd.DataFrame,
	approved: pd.DataFrame,
	db_path: Path,
	base_url: str,
	seed_source_label: str,
) -> None:
	st.header("Painel administrativo")
	if not admin_access(db_path):
		return

	seed_tab, qr_tab, results_tab, voter_stats_tab, audit_tab, deploy_tab = st.tabs(
		["Seed", "QR codes", "Ranking e votos", "Estatísticas de avaliadores", "Auditoria", "Deploy"]
	)
	with seed_tab:
		render_seed_tab(seed, approved, seed_source_label)
	with qr_tab:
		render_qr_tab(approved, base_url)
	with results_tab:
		render_results_tab(db_path)
	with voter_stats_tab:
		render_voter_stats_tab(db_path)
	with audit_tab:
		render_audit_tab(db_path)
	with deploy_tab:
		render_deploy_tab(seed_source_label)


def seed_source_name(seed_source: Any) -> str:
	if hasattr(seed_source, "name"):
		return str(seed_source.name)
	return Path(str(seed_source)).name


def app() -> None:
	view_mode = str(st.query_params.get("view") or "").strip().lower()
	vote_only_view = view_mode == "vote" and requested_work_id() is not None
	vote_exit_view = view_mode == "closed"
	page_config(show_default_heading=not (vote_only_view or vote_exit_view))

	if vote_exit_view:
		render_vote_exit_page()
		return

	default_seed = os.getenv("BEST_WORK_AWARD_SEED_PATH", str(default_seed_path()))

	if vote_only_view:
		st.markdown(
			"""
			<style>
			[data-testid="stSidebar"] {
				display: none;
			}
			</style>
			""",
			unsafe_allow_html=True,
		)
		db_path = DEFAULT_DB_PATH
		sheet_name = "Sheet1"
		seed_source = default_seed
		base_url = os.getenv("BEST_WORK_AWARD_BASE_URL", "")
	else:
		uploaded_seed = st.sidebar.file_uploader(
			"Seed (.csv, .xlsx)",
			type=["csv", "xlsx"],
			help="No deploy em nuvem, prefira CSV versionado no repositório ou um upload administrativo temporário.",
		)
		seed_path = st.sidebar.text_input("Arquivo seed", value=default_seed)
		db_path = Path(st.sidebar.text_input("Banco SQLite", value=str(DEFAULT_DB_PATH)))
		sheet_name = st.sidebar.text_input("Sheet", value="Sheet1")
		base_url = base_url_input()
		seed_source = uploaded_seed if uploaded_seed is not None else seed_path
		if not root_access(db_path):
			return
	seed_source_label = seed_source_name(seed_source)

	try:
		seed = load_seed_data(seed_source, sheet_name)
		approved = approved_seed(seed)
		init_storage(db_path)
	except FileNotFoundError:
		st.error("Arquivo seed não encontrado. Ajuste o caminho ou envie um arquivo pela barra lateral.")
		return
	except ValueError as exc:
		st.error(str(exc))
		return
	except Exception as exc:
		st.error(f"Falha ao iniciar a aplicação: {exc}")
		return

	if vote_only_view:
		render_vote_page(approved, db_path, vote_only=True)
		return

	render_hero(len(approved), storage_backend_label())
	page = st.sidebar.radio("Área", ["Votação", "Admin", "Pódio"], horizontal=True)
	if page == "Votação":
		render_vote_page(approved, db_path, vote_only=False)
	elif page == "Admin":
		render_admin_page(seed, approved, db_path, base_url, seed_source_label)
	else:
		render_podium_page(db_path)


if __name__ == "__main__":
	app()
