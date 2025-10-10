import tkinter as tk
from datetime import datetime, timedelta, time, date

DAGNAVN = ["mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag", "søndag"]

def day_name(d: date) -> str:
    return DAGNAVN[d.weekday()]

def is_weekend(now: datetime) -> bool:
    """Helg = fredag kl. 17:00 → søndag 23:59:59."""
    wd = now.weekday()  # man=0 .. søn=6
    if wd in (5, 6):          # lørdag, søndag
        return True
    if wd == 4 and now.time() >= time(17, 0):  # fredag fra 17:00
        return True
    return False

def days_until_friday(now: datetime) -> int:
    """Hele dager frem til fredag (0 betyr at det er fredag i dag)."""
    return (4 - now.weekday()) % 7

def next_friday_17(now: datetime) -> datetime:
    """Førstkommende fredag kl. 17:00 fra nå."""
    days_to_fri = (4 - now.weekday()) % 7
    candidate = (now + timedelta(days=days_to_fri)).replace(hour=17, minute=0, second=0, microsecond=0)
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate

def fmt_time(dt: datetime) -> str:
    return dt.strftime("%H:%M:%S")

def countdown_text(now: datetime, target: datetime, prefix: str = "") -> str:
    """Returner pen tekst for nedtelling med prefix."""
    delta = target - now
    total_seconds = int(delta.total_seconds())
    if total_seconds < 0:
        total_seconds = 0
    days = total_seconds // 86400
    rem = total_seconds % 86400
    hours = rem // 3600
    minutes = (rem % 3600) // 60
    seconds = rem % 60
    if days > 0:
        txt = f"{prefix}Nedtelling: {days}d {hours:02d}:{minutes:02d}:{seconds:02d} til helg (fre 17:00)."
    else:
        txt = f"{prefix}Nedtelling: {hours:02d}:{minutes:02d}:{seconds:02d} til helg (fre 17:00)."
    return txt

# ------------ GUI ------------
root = tk.Tk()
root.title("Er det fredag?")
root.geometry("520x360")

title = tk.Label(root, text="Er det fredag?", font=("TkDefaultFont", 16, "bold"))
title.pack(pady=(16, 4))

info = tk.Label(root, text="(Helg starter fredag kl. 17:00)", font=("TkDefaultFont", 10))
info.pack(pady=(0, 10))

status_label = tk.Label(root, text="", font=("TkDefaultFont", 13))
status_label.pack(pady=6)

countdown_label = tk.Label(root, text="", font=("TkDefaultFont", 12))
countdown_label.pack(pady=6)

svar_spm = tk.Label(root, text="Sa du 'mandag' nå nettopp? (for ekstra hyggelig beskjed)", font=("TkDefaultFont", 10))
svar_spm.pack(pady=(18, 4))

def start_countdown(target: datetime, prefix: str = ""):
    """Oppdater nedtellingslabel hvert sekund fram til target."""
    def tick():
        now = datetime.now()
        countdown_label.config(text=countdown_text(now, target, prefix))
        if now < target:
            root.after(1000, tick)
        else:
            # Når vi passerer målet: oppdater status til helg.
            status_label.config(text="JA – det er helg! 🎉")
            countdown_label.config(text="")
    tick()

def on_answer(user_says_monday: bool):
    now = datetime.now()

    # Helg: ingen nedtelling, bare hyggelig beskjed
    if is_weekend(now):
        if user_says_monday:
            svar_label.config(text=f"NEI, det er faktisk {day_name(now.date())} – det er fortsatt helg 🎉")
        else:
            svar_label.config(text="Det er fortsatt helg 🎉")
        return

    # Ikke helg → si hvilken dag det er, og start nedtelling til fre 17:00
    n_days = days_until_friday(now)
    prefix = f"NEI, det er faktisk {day_name(now.date())} – " if user_says_monday and day_name(now.date()) != "mandag" else ""
    if day_name(now.date()) == "mandag" and not user_says_monday:
        prefix = "Jo da, det er faktisk mandag – "

    # Tekst om dager til helg
    if n_days == 0:
        lead = f"{prefix}det er fredag – teller ned til helg … "
    elif n_days == 1:
        lead = f"{prefix}1 dag til helg. "
    else:
        lead = f"{prefix}{n_days} dager til helg. "

    target = next_friday_17(now)
    start_countdown(target, prefix=lead)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

btn_yes = tk.Button(btn_frame, text="Ja", width=12, command=lambda: on_answer(True))
btn_no  = tk.Button(btn_frame, text="Nei", width=12, command=lambda: on_answer(False))
btn_yes.pack(side="left", padx=6)
btn_no.pack(side="left", padx=6)

svar_label = tk.Label(root, text="", font=("TkDefaultFont", 11))
svar_label.pack(pady=(12, 22))

root.mainloop()
