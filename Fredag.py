import os
import ssl
import smtplib
import getpass
import tkinter as tk
from email.message import EmailMessage

# ====== STANDARDVERDIER FOR OUTLOOK / O365 ======
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
EMAIL_USER  = os.getenv("EMAIL_USER", "")         # f.eks. navn@firma.no
EMAIL_PASS  = os.getenv("EMAIL_PASS", "")         # passord eller app-passord
EMAIL_TO    = os.getenv("EMAIL_TO", EMAIL_USER)   # send til deg selv hvis ikke satt

def hent_brukernavn() -> str:
    navn = os.getenv("USERNAME") or os.getenv("USER") or getpass.getuser() or "venn"
    for ch in "._-":
        navn = navn.replace(ch, " ")
    navn = " ".join(navn.split()).title()
    return navn or "Venn"

BRUKERNAVN = hent_brukernavn()

def send_epost(emne: str, tekst: str) -> None:
    if not (SMTP_SERVER and SMTP_PORT and EMAIL_USER and EMAIL_PASS and EMAIL_TO):
        svar_label.config(text=f"{tekst}\n(Mangler SMTP-oppsett: EMAIL_USER/PASS/TO)")
        return

    msg = EmailMessage()
    msg["From"] = EMAIL_USER
    msg["To"] = EMAIL_TO
    msg["Subject"] = emne
    msg.set_content(tekst)

    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20) as server:
            server.ehlo()
            server.starttls(context=context)  # STARTTLS for O365
            server.ehlo()
            server.login(EMAIL_USER, EMAIL_PASS)  # må være hele e-postadressen
            server.send_message(msg)
        svar_label.config(text=f"{tekst}\n(E-post sendt til {EMAIL_TO})")
    except Exception as e:
        svar_label.config(text=f"{tekst}\n(Klarte ikke å sende e-post: {e})")

def svar(er_ja: bool) -> None:
    if er_ja:
        tekst = f"god helg, {BRUKERNAVN}!"
    else:
        tekst = f"jo det er helg nå, {BRUKERNAVN}!"
    send_epost("Helg-sjekk", tekst)

# ====== TKINTER-GUI ======
root = tk.Tk()
root.title("Er det snart helg?")

spm_label = tk.Label(root, text="Er det snart helg?")
spm_label.pack(padx=20, pady=(20, 10))

knapperamme = tk.Frame(root)
knapperamme.pack(pady=10)

btn_ja = tk.Button(knapperamme, text="Ja", width=10, command=lambda: svar(True))
btn_nei = tk.Button(knapperamme, text="Nei", width=10, command=lambda: svar(False))
btn_ja.pack(side="left", padx=5)
btn_nei.pack(side="left", padx=5)

svar_label = tk.Label(root, text="", font=("TkDefaultFont", 11, "bold"), justify="center")
svar_label.pack(pady=(10, 20))

root.mainloop()
