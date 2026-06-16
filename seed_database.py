import sqlite3
import random
from datetime import datetime, timedelta

PAGES = ["/home", "/about", "/products", "/blog", "/contact", "/pricing", "/login", "/signup"]
COUNTRIES = ["United States", "United Kingdom", "Canada", "Germany", "France", "Australia", "Brazil", "Mexico"]
SOURCES = ["organic", "direct", "social", "email", "referral", "paid"]

def seed():
    conn = sqlite3.connect("traffic.db")
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS traffic")
    cur.execute("""
        CREATE TABLE traffic (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            page TEXT,
            sessions INTEGER,
            page_views INTEGER,
            bounce_rate REAL,
            avg_duration_sec INTEGER,
            country TEXT,
            source TEXT
        )
    """)

    today = datetime.today()
    rows = []
    for days_back in range(90):
        date = (today - timedelta(days=days_back)).strftime("%Y-%m-%d")
        for page in PAGES:
            for country in random.sample(COUNTRIES, k=random.randint(3, 6)):
                sessions = random.randint(20, 800)
                rows.append((
                    date,
                    page,
                    sessions,
                    int(sessions * random.uniform(1.1, 3.5)),
                    round(random.uniform(0.25, 0.75), 2),
                    random.randint(45, 420),
                    country,
                    random.choice(SOURCES)
                ))

    cur.executemany("""
        INSERT INTO traffic (date, page, sessions, page_views, bounce_rate, avg_duration_sec, country, source)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)

    conn.commit()
    conn.close()
    print(f"Seeded {len(rows)} rows into traffic.db")

if __name__ == "__main__":
    seed()
