from pathlib import Path

import pandas as pd
import requests
import urllib3


def save_photos_from_excel(excel_file: str = "attendees_saf26.xlsx", output_dir: str = "photos") -> None:
	"""Read URL/CHAVE columns from an Excel file and save images as CHAVE.jpg."""
	script_dir = Path(__file__).resolve().parent
	excel_path = script_dir / excel_file
	photos_path = script_dir / output_dir

	if not excel_path.exists():
		raise FileNotFoundError(f"Excel file not found: {excel_path}")

	df = pd.read_excel(excel_path)

	required_columns = {"URL", "CHAVE"}
	missing_columns = required_columns.difference(df.columns)
	if missing_columns:
		missing = ", ".join(sorted(missing_columns))
		raise ValueError(f"Missing required column(s): {missing}")

	photos_path.mkdir(parents=True, exist_ok=True)
	# Petrobras internal endpoints may use SSL interception/self-signed chains.
	urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
	session = requests.Session()
	session.headers.update({"User-Agent": "Mozilla/5.0 (photo-downloader)"})

	saved = 0
	skipped = 0

	for row_number, (_, row) in enumerate(df.iterrows(), start=2):
		url = row.get("URL")
		chave = row.get("CHAVE")

		if pd.isna(url) or str(url).strip() == "":
			skipped += 1
			print(f"[SKIP] Row {row_number}: empty URL")
			continue

		url_text = str(url).strip()
		if not url_text.lower().startswith(("http://", "https://")):
			skipped += 1
			print(f"[SKIP] Row {row_number}: invalid URL '{url_text}'")
			continue

		if pd.isna(chave) or str(chave).strip() == "":
			skipped += 1
			print(f"[SKIP] Row {row_number}: empty CHAVE")
			continue

		filename = f"{str(chave).strip()}.jpg"
		destination = photos_path / filename

		try:
			response = session.get(url_text, timeout=(8, 20), verify=True)
			response.raise_for_status()
			destination.write_bytes(response.content)
			saved += 1
			print(f"[OK] {destination.name}")
		except requests.exceptions.SSLError:
			try:
				response = session.get(url_text, timeout=(8, 20), verify=False)
				response.raise_for_status()
				destination.write_bytes(response.content)
				saved += 1
				print(f"[OK] {destination.name} (SSL verify disabled)")
			except requests.RequestException as exc:
				skipped += 1
				print(f"[ERROR] Row {row_number} ({destination.name}): {exc}")
		except requests.RequestException as exc:
			skipped += 1
			print(f"[ERROR] Row {row_number} ({destination.name}): {exc}")
		except KeyboardInterrupt:
			print("\nInterrupted by user. Finishing execution gracefully.")
			break

	print(f"Done. Saved: {saved}, Skipped/Errors: {skipped}")


if __name__ == "__main__":
	try:
		save_photos_from_excel()
	except KeyboardInterrupt:
		print("\nInterrupted by user.")
