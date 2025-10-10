import tkinter as tk
from datetime import datetime, timedelta, time, date

# --- KONFIG ---
WEEKEND_CUTOFF = time(16, 15)  # fredag kl. 16:10
DAGNAVN = ["mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag", "søndag"]

# --------- Tidshjelpere ---------
def day_name(d: date) -> str:
    return DAGNAVN[d.weekday()]

def next_friday_cutoff(now: datetime) -> datetime:
    """Førstkommende fredag kl. WEEKEND_CUTOFF fra 'now'."""
    days_to_fri = (4 - now.weekday()) % 7
    candidate = (now + timedelta(days=days_to_fri)).replace(
        hour=WEEKEND_CUTOFF.hour, minute=WEEKEND_CUTOFF.minute, second=0, microsecond=0
    )
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate

def next_monday_midnight(now: datetime) -> datetime:
    """Neste mandag 00:00 etter 'now'."""
    days_to_mon = (7 - now.weekday()) % 7  # 0 hvis mandag i dag
    candidate = (now + timedelta(days=days_to_mon)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate

def is_weekend(now: datetime) -> bool:
    """Helg = fredag fra WEEKEND_CUTOFF og hele lørdag/søndag."""
    wd = now.weekday()
    if wd in (5, 6):
        return True
    if wd == 4 and now.time() >= WEEKEND_CUTOFF:
        return True
    return False

def days_until_friday(now: datetime) -> int:
    """Hele dager (0–6) til fredag (ignorer klokkeslett)."""
    return (4 - now.weekday()) % 7

# --------- Nedtelling ---------
countdown_job = None

def cancel_countdown():
    global countdown_job
    if countdown_job:
        root.after_cancel(countdown_job)
        countdown_job = None

def start_countdown(target: datetime, prefix: str = ""):
    cancel_countdown()
    update_countdown(target, prefix)

def update_countdown(target: datetime, prefix: str):
    global countdown_job
    now = datetime.now()
    remaining = target - now
    if remaining.total_seconds() <= 0:
        svar_label.config(text="🎉 GOD HELG! 🎉" if is_weekend(datetime.now()) else "Helgen er over. God mandag!")
        countdown_job = None
        return

    total = int(remaining.total_seconds())
    days, rem = divmod(total, 86400)
    hours, rem = divmod(rem, 3600)
    minutes, seconds = divmod(rem, 60)

    if days > 0:
        txt = f"{prefix}{days}d {hours:02d}:{minutes:02d}:{seconds:02d}"
    else:
        txt = f"{prefix}{hours:02d}:{minutes:02d}:{seconds:02d}"

    svar_label.config(text=txt)
    countdown_job = root.after(1000, update_countdown, target, prefix)

# --------- Helgesjekk ---------
def helgesjekk():
    now = datetime.now()
    cancel_countdown()

    if is_weekend(now):
        # Det er helg: tell hvor lenge helgen varer (til mandag 00:00)
        today = day_name(now.date())
        slutt = next_monday_midnight(now)
        prefix = f"Det er {today} – det er helg 🎉  Helgen varer i: "
        start_countdown(slutt, prefix)
    else:
        # Ikke helg: si dag, dager til helg og tell ned til fredag 16:10
        today = day_name(now.date())
        n = days_until_friday(now)
        if n == 0:
            info = "Det er fredag – nedtelling til helg (fre 16:10): "
        elif n == 1:
            info = f"Det er {today} – 1 dag til helg. Nedtelling: "
        else:
            info = f"Det er {today} – {n} dager til helg. Nedtelling: "
        start_countdown(next_friday_cutoff(now), info)

# --------- GUI ---------
root = tk.Tk()
root.title("Helgesjekk (nedtelling til/fra helg)")

tk.Label(root, text="Helgesjekk", font=("TkDefaultFont", 13, "bold")).pack(padx=20, pady=(18, 8))

btn = tk.Button(root, text="Sjekk helg", width=16, command=helgesjekk)
btn.pack(pady=6)

svar_label = tk.Label(root, text="Trykk «Sjekk helg» for status.", font=("TkDefaultFont", 11))
svar_label.pack(pady=(10, 18))

root.mainloop()

