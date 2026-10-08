"""Create demo_inbox/: 36 monthly branch sales files with inconsistent names, mixed in with PDFs, images and notes,
the way a shared "Downloads" or "Reports" folder looks after a year. The seed is fixed so it's the same every run."""
import random
from pathlib import Path

import pandas as pd

random.seed(11)
INBOX = Path(__file__).parent / "demo_inbox"
BRANCHES = ["Cairo", "Alexandria", "Port Said"]
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
PRODUCTS = {"Keyboard": 650, "Wireless Mouse": 350, "Headphones": 1200, "Webcam": 1500, "USB-C Cable": 120}
NAME_STYLES = ["sales_{b}_{m}_2024.xlsx", "{B} {M} 2024 report.xlsx", "{m}-{b}-sales-final.xlsx", "Sales {M}24 ({B}).xlsx"]


def main():
    INBOX.mkdir(exist_ok=True)
    for b in BRANCHES:
        for i, m in enumerate(MONTHS, 1):
            rows = [{"date": f"2024-{i:02d}-{random.randint(1, 28):02d}", "product": p,
                     "quantity": random.randint(1, 20), "unit_price": price}
                    for p, price in PRODUCTS.items() for _ in range(random.randint(2, 6))]
            name = random.choice(NAME_STYLES).format(b=b.lower().replace(" ", ""), B=b, m=m, M=m.title())
            pd.DataFrame(rows).to_excel(INBOX / name, index=False)
    for n in range(8):
        (INBOX / f"invoice_{1000 + n}.pdf").write_bytes(b"%PDF-1.4\n% demo invoice\n")
        (INBOX / f"IMG_20240{n + 1}1{n}_1200{n}.jpg").write_bytes(b"\xff\xd8\xff demo image")
    for n in ("notes.txt", "todo - old.txt", "meeting notes march.txt"):
        (INBOX / n).write_text("demo note\n", encoding="utf-8")
    print(f"{sum(1 for _ in INBOX.iterdir())} files in {INBOX}")


if __name__ == "__main__":
    main()
