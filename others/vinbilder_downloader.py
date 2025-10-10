# vinbilder_downloader_v3.py
import os, pathlib, re, time, requests, json, html

VNR = [
    11253001, 17191001, 14472901, 15075201, 15828401,
    15075101, 15913701, 10100001, 13017001, 14675301,
    13572601, 17190601, 12035301, 15726502, 15183901,
    17123401, 19215501, 17370707
]

SAVE = r"C:\Users\ib91\Desktop\Vinmonopolet\Bilder viner"
UA = {"User-Agent": "Mozilla/5.0 (BildeFetcher v2)"}
TIMEOUT = 10
pathlib.Path(SAVE).mkdir(parents=True, exist_ok=True)

def get_real_photo_id(varnr: int) -> str | None:
    url = f"https://www.vinmonopolet.no/p/{varnr}"
    try:
        html_text = requests.get(url, headers=UA, timeout=TIMEOUT).text
    except requests.RequestException:
        return None
    # prøv å finne "productPhoto":"123456.jpg"
    m = re.search(r'"productPhoto"\s*:\s*"([^"]+\.jpg)"', html_text)
    if m:
        return m.group(1)
    return None

def download(url: str, dest: str) -> bool:
    try:
        r = requests.get(url, headers=UA, timeout=TIMEOUT)
        if r.status_code == 200 and r.headers.get("content-type","").startswith("image"):
            with open(dest, "wb") as f:
                f.write(r.content)
            return True
    except requests.RequestException:
        pass
    return False

for varnr in VNR:
    photo = get_real_photo_id(varnr)
    if not photo:
        print(f"✖ Fant ikke productPhoto for {varnr}")
        continue

    img_url = f"https://bilder.vinmonopolet.no/cache/960x960-0/{photo}"
    dest    = os.path.join(SAVE, f"{varnr}.jpg")

    if download(img_url, dest):
        print(f"✔ Lagret {dest}")
    else:
        print(f"✖ Feil nedlasting {varnr} ({img_url})")

    time.sleep(0.4)

print("\nFerdig!")
