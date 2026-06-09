import os
import sqlite3
import unittest
from datetime import datetime
import pandas as pd

import db_manager


class TestBarcodeRoutingLogic(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        db_manager.DB_FILE             = 'test_app.db'
        db_manager.CODE_LOOKUP_EXCEL   = 'test_code_lookup.xlsx'
        db_manager.ROUTING_RECORD_EXCEL = 'test_routing_record.xlsx'

    @classmethod
    def tearDownClass(cls):
        for f in ['test_app.db', 'test_code_lookup.xlsx', 'test_routing_record.xlsx']:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except Exception as e:
                print(f"Cleanup warning: {e}")

    def setUp(self):
        # Delete database and recreate a fresh schema on every test to prevent schema pollution
        if os.path.exists(db_manager.DB_FILE):
            try:
                os.remove(db_manager.DB_FILE)
            except Exception:
                pass
        db_manager.init_db()

    # ─────────────────────────────────────────────────────────────────────────
    def test_database_initialization(self):
        """Verify SQLite tables are created with the correct schema."""
        conn   = db_manager.get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='code_lookup'")
        self.assertIsNotNone(cursor.fetchone(), "code_lookup table missing")

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='routing_records'")
        self.assertIsNotNone(cursor.fetchone(), "routing_records table missing")

        # Verify receiver_name columns exist
        cursor.execute("PRAGMA table_info(routing_records)")
        cols = [row[1] for row in cursor.fetchall()]
        self.assertIn('receiver_name_1', cols, "receiver_name_1 column missing")
        self.assertIn('receiver_name_10', cols, "receiver_name_10 column missing")
        conn.close()

    # ─────────────────────────────────────────────────────────────────────────
    def test_routing_multi_row_expansion_and_stages(self):
        """Scanning a 3-employee barcode splits into 3 rows and stages route correctly."""
        code      = "TEST01"
        dept      = "SDO - MANILA"
        school    = "Riverside Elementary"
        employees = ["Alice Doe", "Bob Ross", "Charlie Smith"]
        doc_type  = "Clearance"

        conn = db_manager.get_db_connection()
        conn.execute(
            "INSERT INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (code, dept, school, ",".join(employees), doc_type, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()

        ts1 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for emp in employees:
            row = conn.execute(
                "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
            ).fetchone()
            self.assertIsNone(row)
            conn.execute(
                "INSERT INTO routing_records "
                "(department, school_office, employee, code, receiving_office_1, receiver_name_1, timestamp_1) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (dept, school, emp, code, "Office Alpha", "John Doe", ts1)
            )
        conn.commit()

        rows = conn.execute("SELECT * FROM routing_records WHERE code=?", (code,)).fetchall()
        self.assertEqual(len(rows), 3)
        for r in rows:
            self.assertEqual(r['receiving_office_1'], "Office Alpha")
            self.assertEqual(r['receiver_name_1'],    "John Doe")
            self.assertIsNotNone(r['timestamp_1'])
            self.assertIsNone(r['receiving_office_2'])

        # Duplicate prevention simulation
        latest_offices = []
        for r in rows:
            for i in range(10, 0, -1):
                if r[f'receiving_office_{i}']:
                    latest_offices.append(r[f'receiving_office_{i}'])
                    break
        self.assertTrue(all(o == "Office Alpha" for o in latest_offices))

        # Advance to stage 2
        ts2 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for emp in employees:
            r = conn.execute(
                "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
            ).fetchone()
            conn.execute(
                "UPDATE routing_records SET receiving_office_2=?, receiver_name_2=?, timestamp_2=? WHERE id=?",
                ("Office Beta", "Jane Smith", ts2, r['id'])
            )
        conn.commit()

        updated = conn.execute("SELECT * FROM routing_records WHERE code=?", (code,)).fetchall()
        for r in updated:
            self.assertEqual(r['receiving_office_1'], "Office Alpha")
            self.assertEqual(r['receiving_office_2'], "Office Beta")
            self.assertEqual(r['receiver_name_2'],    "Jane Smith")
            self.assertIsNone(r['receiving_office_3'])
        conn.close()

    # ─────────────────────────────────────────────────────────────────────────
    def test_dynamic_expansion_beyond_10_stages(self):
        """Scanning a barcode 12× at different offices auto-expands to stage 12."""
        code   = "EXPAND1"
        dept   = "Elementary Schools"
        school = "Northside Elementary"
        emp    = "Dynamic Test Employee"

        conn = db_manager.get_db_connection()

        # Register the code
        conn.execute(
            "INSERT INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (code, dept, school, emp, "Service Record", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()

        # Simulate 12 sequential scans at different offices
        offices = [f"Office {chr(65 + i)}" for i in range(12)]  # Office A through Office L

        for idx, office in enumerate(offices, start=1):
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            row = conn.execute(
                "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
            ).fetchone()

            if not row:
                # First scan
                db_manager.ensure_routing_columns(conn, 1)
                conn.execute(
                    "INSERT INTO routing_records "
                    "(department, school_office, employee, code, receiving_office_1, receiver_name_1, timestamp_1) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (dept, school, emp, code, office, "Test Receiver", ts)
                )
                conn.commit()
            else:
                # Find next empty slot, expanding schema if needed
                office_cols = [c for c in row.keys() if c.startswith('receiving_office_')]
                stages      = sorted({int(c.split('_')[-1]) for c in office_cols if c.split('_')[-1].isdigit()})
                if not stages:
                    stages = list(range(1, 11))

                empty_slot = None
                for s in stages:
                    k = f'receiving_office_{s}'
                    if k in row.keys() and not row[k]:
                        empty_slot = s
                        break

                if empty_slot is None:
                    next_slot = max(stages) + 1
                    db_manager.ensure_routing_columns(conn, next_slot)
                    empty_slot = next_slot

                conn.execute(
                    f"UPDATE routing_records "
                    f"SET receiving_office_{empty_slot}=?, receiver_name_{empty_slot}=?, timestamp_{empty_slot}=? "
                    f"WHERE id=?",
                    (office, "Test Receiver", ts, row['id'])
                )
                conn.commit()

                # Re-fetch updated row for next iteration
                row = conn.execute(
                    "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
                ).fetchone()

        # Verify stages 1–12 all populated
        final_row = conn.execute(
            "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
        ).fetchone()
        self.assertIsNotNone(final_row, "Employee row missing after 12 scans")

        for i in range(1, 13):
            key = f'receiving_office_{i}'
            self.assertIn(key, final_row.keys(), f"Column {key} missing from schema")
            self.assertEqual(final_row[key], offices[i - 1], f"Stage {i} office mismatch")

        # Confirm columns 11 and 12 were dynamically added
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(routing_records)")
        col_names = [row[1] for row in cursor.fetchall()]
        self.assertIn('receiving_office_11', col_names, "Stage 11 column not auto-created")
        self.assertIn('receiver_name_11',    col_names, "receiver_name_11 not auto-created")
        self.assertIn('timestamp_11',        col_names, "timestamp_11 not auto-created")
        self.assertIn('receiving_office_12', col_names, "Stage 12 column not auto-created")
        self.assertIn('receiver_name_12',    col_names, "receiver_name_12 not auto-created")
        self.assertIn('timestamp_12',        col_names, "timestamp_12 not auto-created")

        conn.close()
        print("Dynamic expansion beyond 10 stages: PASSED (12 offices tracked)")

    # ─────────────────────────────────────────────────────────────────────────
    def test_ensure_routing_columns_idempotent(self):
        """Calling ensure_routing_columns multiple times does not raise errors."""
        conn = db_manager.get_db_connection()
        db_manager.ensure_routing_columns(conn, 15)
        db_manager.ensure_routing_columns(conn, 15)  # Should not raise
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(routing_records)")
        cols = [row[1] for row in cursor.fetchall()]
        self.assertIn('receiving_office_15', cols)
        self.assertIn('receiver_name_15',    cols)
        self.assertIn('timestamp_15',        cols)
        conn.close()

    # ─────────────────────────────────────────────────────────────────────────
    def test_receiver_name_stored_correctly(self):
        """Receiver name is stored alongside the office in each stage slot."""
        code  = "RCVR01"
        dept  = "Secondary Schools"
        school = "Westside High"
        emp   = "Receiver Test Employee"

        conn = db_manager.get_db_connection()
        conn.execute(
            "INSERT INTO code_lookup (code, department, school_office, employees, doc_type, generated_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (code, dept, school, emp, "Transfer Document", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()

        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db_manager.ensure_routing_columns(conn, 1)
        conn.execute(
            "INSERT INTO routing_records "
            "(department, school_office, employee, code, receiving_office_1, receiver_name_1, timestamp_1) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (dept, school, emp, code, "Personnel Office", "Maria Santos", ts)
        )
        conn.commit()

        row = conn.execute(
            "SELECT * FROM routing_records WHERE code=? AND employee=?", (code, emp)
        ).fetchone()
        self.assertEqual(row['receiving_office_1'], "Personnel Office")
        self.assertEqual(row['receiver_name_1'],    "Maria Santos")
        self.assertIsNotNone(row['timestamp_1'])
        conn.close()


if __name__ == '__main__':
    unittest.main(verbosity=2)
