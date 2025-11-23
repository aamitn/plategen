# NAMEPLATE LIST EXCEL GENERATOR - DATABASE SCHEMA GENERATOR
import sqlite3

# Database file
DB_FILE = 'nameplates.db'

def create_tables(conn):
    cursor = conn.cursor()

    # Plate Types Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS plate_types (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type_name TEXT NOT NULL,
        default_size TEXT
    )
    ''')

    # Charger Groups Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS ch_groups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_name TEXT NOT NULL
    )
    ''')

    # Nameplates Table with repeater field
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS nameplates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sl_no INTEGER NOT NULL,
        type_id INTEGER NOT NULL,
        ch_group_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        qty INTEGER NOT NULL,
        repeater INTEGER DEFAULT 0,
        FOREIGN KEY (type_id) REFERENCES plate_types(id) ON DELETE RESTRICT ON UPDATE CASCADE,
        FOREIGN KEY (ch_group_id) REFERENCES ch_groups(id) ON DELETE RESTRICT ON UPDATE CASCADE
    )
    ''')

    conn.commit()
    print("Tables created successfully!")

def insert_default_data(conn):
    cursor = conn.cursor()

    # Insert plate types with updated sizes
    plate_types = [
        ('Rectangular', '75x15'),
        ('Rectangular', '40x15'),
        ('Rectangular', '25x15'),
        ('Ring', '22Φ'),
        ('Ring', '14Φ'),
        ('Ring', '12Φ')
    ]
    cursor.executemany('INSERT INTO plate_types (type_name, default_size) VALUES (?, ?)', plate_types)

    # Insert charger groups
    ch_groups = [
        ('SFCB',),
        ('DFCB',),
        ('FFCB',),
        ('COMMON',)
    ]
    cursor.executemany('INSERT INTO ch_groups (group_name) VALUES (?)', ch_groups)

    conn.commit()
    print("Default plate types and charger groups inserted!")

def insert_sample_nameplates(conn):
    cursor = conn.cursor()

    # Insert sample nameplates including repeater field (0 = no repeat)
    nameplates = [
        (1, 1, 1, 'CHARGER PANEL', 2, 0),
        (2, 4, 2, 'RING BUTTON', 5, 0),
        (3, 2, 3, 'RECTANGULAR DISPLAY', 3, 1),  # repeater = 1
        (4, 5, 4, 'RING INDICATOR', 4, 2),       # repeater = 2
        (5, 3, 1, 'LARGE RECTANGULAR', 1, 0)
    ]

    cursor.executemany('''
        INSERT INTO nameplates (sl_no, type_id, ch_group_id, name, qty, repeater)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', nameplates)

    conn.commit()
    print("Sample nameplate entries inserted!")

def main():
    # Connect to SQLite DB (creates if not exists)
    conn = sqlite3.connect(DB_FILE)
    print(f"Connected to database '{DB_FILE}'")

    # Create tables
    create_tables(conn)

    # Insert default plate types and charger groups
    insert_default_data(conn)

    # Insert sample nameplates
    insert_sample_nameplates(conn)

    # Close connection
    conn.close()
    print("Database setup complete!")

if __name__ == '__main__':
    main()
