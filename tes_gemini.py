"""Diagnosis koneksi Gemini - jalankan SEBELUM memakai aplikasi.

    python tes_gemini.py                 (kunci dibaca dari GEMINI_API_KEY, atau diminta)
    python tes_gemini.py nama-model ...  (opsional: uji model tertentu)

Hasilnya memberi tahu salah satu dari:
  - kunci bermasalah (401/403)         -> buat kunci baru di Google AI Studio
  - model tidak tersedia (404)         -> pakai nama model dari daftar di bawah
  - kuota habis (429 / limit: 0)       -> tunggu reset, atau aktifkan billing
  - semua OK                           -> isi GEMINI_MODEL dengan model yang lolos
"""
import os
import sys
import time

try:
    import google.genai as genai_pkg
    from google import genai
except ImportError:
    sys.exit("Library belum terpasang. Jalankan:  pip install -U google-genai")

print("google-genai versi:", getattr(genai_pkg, "__version__", "tidak diketahui"))

key = (os.getenv("GEMINI_API_KEY") or input("Tempel API key: ")).strip()
print(f"Kunci: awalan '{key[:3]}', panjang {len(key)} karakter "
      "(kunci AQ. umumnya panjang; jika terlalu pendek, kemungkinan terpotong saat disalin)")

client = genai.Client(api_key=key)   # semua format kunci lewat api_key, tanpa header Bearer

print("\n1) Model yang bisa dipakai kunci ini:")
try:
    names = []
    for m in client.models.list():
        actions = getattr(m, "supported_actions", None) or []
        if "generateContent" in actions:
            names.append(m.name.replace("models/", ""))
    names = sorted(set(names))
    for n in names:
        if n.startswith("gemini"):
            print("   ", n)
except Exception as e:
    print("   GAGAL:", type(e).__name__, str(e)[:500])
    print("\n=> Masalah ada di KUNCI/autentikasi (bukan di model).")
    print("   Buat kunci baru di Google AI Studio, salin UTUH, lalu coba lagi.")
    print("   Pastikan juga: pip install -U google-genai")
    sys.exit(1)

skip = ("image", "tts", "live", "audio", "embedding", "robotics", "preview", "native")
auto = [n for n in names if n.startswith("gemini-3") and "flash" in n and not any(x in n for x in skip)]
candidates = sys.argv[1:] or sorted(auto, reverse=True)[:3] or ["gemini-3.5-flash"]

print("\n2) Uji panggilan:")
good = []
for model in candidates:
    t0 = time.time()
    try:
        r = client.models.generate_content(model=model, contents="Balas hanya dengan kata: OK")
        print(f"   [OK]    {model}  ({time.time() - t0:.1f} dtk) -> {(r.text or '').strip()[:30]!r}")
        good.append(model)
    except Exception as e:
        print(f"   [GAGAL] {model} -> {type(e).__name__}: {str(e)[:350]}")

print()
if good:
    print(f"=> Gunakan model ini di aplikasi:  GEMINI_MODEL = \"{good[0]}\"")
else:
    print("=> Semua gagal. Baca pesan error di atas:")
    print("   401/403 = kunci;  404 = model tidak ada;  429 = kuota (cek 'limit: 0' / 'PerDay').")
