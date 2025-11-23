-- import to SQLite by running: sqlite3.exe db.sqlite3 -init sqlite.sql

PRAGMA journal_mode = MEMORY;
PRAGMA synchronous = OFF;
PRAGMA foreign_keys = OFF;
PRAGMA ignore_check_constraints = OFF;
PRAGMA auto_vacuum = NONE;
PRAGMA secure_delete = OFF;
BEGIN TRANSACTION;

;
CREATE TABLE IF NOT EXISTS ch_groups (
id INTEGER PRIMARY KEY AUTOINCREMENT,
group_name TEXT NOT NULL
);
DELETE FROM "ch_groups";
INSERT INTO "ch_groups" ("id", "group_name") VALUES
(1, 'SFCB'),
(2, 'DFCB'),
(3, 'FFCB'),
(4, 'COMMON'),
(5, 'SPECIAL'),
(6, 'SFCB'),
(7, 'DFCB'),
(8, 'FFCB'),
(9, 'COMMON');

CREATE TABLE IF NOT EXISTS "nameplates" (
id INTEGER PRIMARY KEY AUTOINCREMENT,
sl_no INTEGER NOT NULL,
type_id INTEGER NOT NULL,
ch_group_id INTEGER NOT NULL,
name TEXT NOT NULL,
qty INTEGER NOT NULL,
repeater INTEGER DEFAULT 0,
FOREIGN KEY (type_id) REFERENCES plate_types(id) ON DELETE RESTRICT ON UPDATE CASCADE,
FOREIGN KEY (ch_group_id) REFERENCES ch_groups(id) ON DELETE RESTRICT ON UPDATE CASCADE
);
DELETE FROM "nameplates";
INSERT INTO "nameplates" ("id", "sl_no", "type_id", "ch_group_id", "name", "qty", "repeater") VALUES
(1, 1, 5, 4, 'TEST', 2, 0),
(2, 2, 5, 4, 'ACCEPT', 5, 0),
(3, 3, 5, 4, 'RESET', 3, 0),
(4, 4, 5, 4, 'OCR RESET', 4, 1),
(5, 5, 5, 4, 'OVR RESET', 1, 1),
(6, 6, 6, 4, 'MANUAL VOLTAGE ADJ.', 2, 0),
(7, 7, 2, 4, 'HOOTER', 2, 0),
(8, 8, 3, 4, 'SOCKET', 2, 0),
(9, 9, 3, 4, 'PANEL LAMP', 2, 0),
(10, 10, 3, 4, 'MANUAL BOOST ON', 2, 1),
(11, 11, 1, 4, 'INPUT MCB', 2, 1),
(12, 12, 1, 4, 'AUTO MANUAL SELECTOR SWITCH (2=MANUAL,1+3=OFF,4=AUTO)', 2, 1),
(13, 13, 4, 4, 'AC SUPPLY UNHEALTHY', 2, 0),
(14, 14, 4, 4, 'SUPPLY ON (R)', 2, 0),
(15, 15, 4, 4, 'SUPPLY ON (Y)', 2, 0),
(16, 16, 4, 4, 'SUPPLY ON (B)', 2, 0),
(17, 17, 4, 4, 'OUTPUT ON (FLOAT)', 2, 0),
(18, 18, 4, 4, 'OUTPUT ON (BOOST)', 2, 0),
(19, 19, 4, 4, 'DC OVER CURRENT', 2, 1),
(20, 19, 4, 4, 'CHARGER FAIL', 2, 1),
(21, 19, 4, 4, 'LOAD OVER VOLTAGE', 2, 0),
(22, 19, 4, 4, 'LOAD UNDER VOLTAGE', 2, 0),
(23, 23, 4, 4, 'BATTERY  EARTH FAULT', 2, 0),
(24, 24, 4, 4, 'RECTIFIER FUSE FAIL', 2, 2),
(25, 25, 4, 4, 'FILTER CAP. FUSE FAIL', 2, 2),
(26, 26, 1, 2, 'CHARGER - I', 2, 0),
(27, 27, 1, 2, 'CHARGER - II', 2, 0),
(28, 27, 1, 4, 'INPUT VOLTMETER ', 2, 0),
(29, 27, 1, 4, 'INPUT VOLTMETER SELECTOR SWITCH   ', 2, 0),
(30, 27, 1, 2, 'OUTPUT VOLTMETER (CH-I/CH-II/LOAD/BATT)', 2, 0),
(31, 27, 1, 3, 'OUTPUT VOLTMETER (FC/FCB/LOAD/BATT)', 2, 0),
(32, 27, 1, 4, 'INPUT AMMETER', 2, 0),
(33, 33, 1, 4, 'INPUT AMMETER SELECTOR SWITCH', 2, 0),
(34, 34, 1, 4, 'OUTPUT AMMETER', 2, 2),
(35, 34, 1, 1, 'OUTPUT VOLTMETER (CH/BATTERY)', 2, 1),
(36, 34, 1, 1, 'OUTPUT VOLTAGE SELECTOR SWITCH  (CHARGER/BATT.)', 2, 1),
(37, 34, 1, 1, 'BATTERY CHARGE/DISCHARGE AMMETER', 2, 1),
(38, 34, 1, 5, 'SPECIAL', 2, 1),
(39, 1, 1, 1, 'CHARGER PANEL', 2, 0),
(40, 2, 4, 2, 'RING BUTTON', 5, 0),
(41, 3, 2, 3, 'RECTANGULAR DISPLAY', 3, 1),
(42, 4, 5, 4, 'RING INDICATOR', 4, 2),
(43, 5, 3, 1, 'LARGE RECTANGULAR', 1, 0);
CREATE TABLE IF NOT EXISTS plate_types (
id INTEGER PRIMARY KEY AUTOINCREMENT,
type_name TEXT NOT NULL,
default_size TEXT
);
DELETE FROM "plate_types";
INSERT INTO "plate_types" ("id", "type_name", "default_size") VALUES
(1, 'Rectangular', '75x15'),
(2, 'Rectangular', '4X'15''),
(3, 'Rectangular', '25x15'),
(4, 'Ring', '22Φ'),
(5, 'Ring', '14Φ'),
(6, 'Ring', '12Φ'),
(7, 'Rectangular', '75x15'),
(8, 'Rectangular', '4X'15''),
(9, 'Rectangular', '25x15'),
(10, 'Ring', '22Φ'),
(11, 'Ring', '14Φ'),
(12, 'Ring', '12Φ');





COMMIT;
PRAGMA ignore_check_constraints = ON;
PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
