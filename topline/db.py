import sqlite3
import json
from pathlib import Path
import logging
import csv
from topline.transaction import Transaction

logger = logging.getLogger(__name__)


class DB:
    def __init__(self, filename='topline.db', user_file=None, transaction_file=None):
        self.filename = filename
        if not Path(filename).is_file() and not user_file:
            logger.error('Database and User file not found. Unable to create database')
            self.connection = None
        else:
            logger.info("Connecting to database %s", self.filename)
            self.connection = sqlite3.connect(self.filename, detect_types=sqlite3.PARSE_DECLTYPES)
            self.cursor = self.connection.cursor()
            self.tables = self.get_tables()
            if ('users' not in self.tables or not self.get_usernames()) and not user_file:
                logger.error("Unable to initialise users table without user json file")
                self.close_db()
            else:
                self.initialise_db(user_file, transaction_file)

    def initialise_db(self, user_file=None, transactions_file=None):
        logger.info("Initialising database tables")
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS users(
                                            id INTEGER PRIMARY KEY,
                                            name TEXT NOT NULL, 
                                            surname TEXT NOT NULL,
                                            email TEXT UNIQUE NOT NULL,
                                            username TEXT UNIQUE NOT NULL,
                                            alt_username TEXT,
                                            last_transaction_id INTEGER,
                                            total REAL,
                                            share REAL)
                            ''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS accounts(
                                            acc_num INTEGER PRIMARY KEY,
                                            name TEXT NOT NULL,
                                            balance REAL
                                            active INTEGER NOT NULL)
                            ''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS transactions(
                                            id INTEGER PRIMARY KEY,
                                            acc_num INTEGER NOT NULL,
                                            date DATE NOT NULL,
                                            description TEXT,
                                            reference TEXT,
                                            amount REAL NOT NULL,
                                            user_id INTEGER,
                                            contrib_month TEXT,
                                            contrib_year INTEGER)
                            ''')
        self.cursor.execute('''CREATE UNIQUE INDEX IF NOT EXISTS unique_transaction 
                                            ON transactions(acc_num, date, description, reference, amount)
                            ''')
        self.cursor.execute('''CREATE TRIGGER IF NOT EXISTS calc_share AFTER UPDATE OF total ON users
                               BEGIN
                                    UPDATE users
                                    SET share = (total / (SELECT SUM(total) FROM users WHERE username <> 'TIG') * 100) 
                                    WHERE username <> 'TIG';
                               END
                            ''')

        self.connection.commit()
        if user_file:
            self.initialise_users(user_file)
        if transactions_file:
            self.initialise_transactions(transactions_file)

    def get_tables(self):
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        return [name[0] for name in self.cursor.fetchall()]

    @staticmethod
    def get_users_from_json(json_filename):
        logger.info("Reading user json file %s", json_filename)
        with open(json_filename, 'r') as f:
            users = json.load(f)
        return users

    def initialise_users(self, filename):
        logger.info("Initialising users table")
        users = self.get_users_from_json(filename)
        for key, value in users.items():
            if 'alt_id' in value:
                key2 = value['alt_id']
            else:
                key2 = None
            logger.debug("Adding user: name = %s, surname = %s, email = %s, username = %s",
                         value['firstName'], value['lastName'], value['email'], key)
            try:
                self.cursor.execute(
                    "INSERT INTO users (name, surname, email, username, alt_username) VALUES (?,?,?,?,?)",
                    (value['firstName'], value['lastName'], value['email'], key, key2))
            except sqlite3.IntegrityError:
                logger.warning("User %s %s already in database", value['firstName'], value['lastName'])
        self.connection.commit()

    def initialise_transactions(self, filename):
        logger.info("Initialising transactions table")
        if not Path(filename).is_file():
            logger.warning('Transaction history file not found: %s. Unable to initialise transactions table',
                           Path(filename).absolute())
        else:
            transactions = self.get_transactions_from_csv(filename)
            logger.info("Read %s transactions from file", len(transactions))
            count = 0
            Transaction.usernames = self.get_usernames()
            Transaction.accounts = self.get_accounts()
            for trans in transactions:
                logger.debug("Processing transaction %s/%s: date = %s, desc = %s, ref = %s, amount = %s.",
                             count + 1, len(transactions), trans[0][0], trans[0][1], trans[0][2], trans[0][4])
                t = Transaction(trans[0], trans[1])
                if self.check_transaction(t.account, t.date, t.description, t.reference, t.amount):
                    continue
                contrib = t.process_transaction()
                t.transaction_id = self.add_transaction(t.account, t.date, t.description, t.reference, t.amount,
                                                        t.user_id, t.month, t.year)
                if contrib and t.transaction_id:
                    self.update_user(t.user_id, t.username, t.date, t.amount, t.transaction_id)
                count += 1
            logger.info("Processed %s/%s transactions in %s.", count, len(transactions), filename)

    @staticmethod
    def get_transactions_from_csv(filename):
        logger.info("Reading transactions from %s file", filename)
        transactions = []
        with open(filename, 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                transactions.append([[row[0].replace('-', ' '), row[1], row[2], '', row[3], row[4]], int(row[5])])
        return transactions

    def get_usernames(self):
        logger.debug("Fetching usernames from database")
        result = self.cursor.execute("SELECT id, username, alt_username FROM users ORDER BY id")
        user_list = result.fetchall()
        if not user_list:
            return False
        logger.info("Fetched %s users from database", len(user_list))
        return user_list

    def get_users(self):
        logger.debug("Fetching users from database")
        result = self.cursor.execute('''SELECT u.id, u.name, u.surname, u.email, u.username, t.date, t.amount,
                                               t.contrib_month, t.contrib_year, u.total, u.share
                                        FROM users u LEFT JOIN transactions t
                                        ON u.last_transaction_id=t.id
                                     ''')
        user_list = result.fetchall()
        if not user_list:
            return False
        logger.info("Fetched %s users from database", len(user_list))
        return user_list

    def update_user(self, user_id, username, date, amount, transaction_id):
        result = self.cursor.execute("SELECT total FROM users WHERE id = ?", (user_id,))
        total = result.fetchone()[0] or 0
        share = None
        logger.info("Updating user %s: R %.2f received on %s, total: R %.2f -> R %.2f",
                    username, amount, date, total, total + amount)
        self.cursor.execute('''UPDATE users 
                               SET last_transaction_id = ?,
                                   total = ?,
                                   share = ?
                               WHERE id = ?''',
                            (transaction_id, total + amount, share, user_id))
        self.connection.commit()

    def get_accounts(self):
        logger.debug("Fetching accounts from database")
        result = self.cursor.execute("SELECT acc_num, name, active FROM accounts")
        accounts = result.fetchall()
        logger.info("Fetched %s accounts from database", len(accounts))
        return accounts

    def add_account(self, acc_num, name, balance, active=True):
        logger.info("Adding account to database: %s - %s, balance = R %.2f", acc_num, name, float(balance))
        self.cursor.execute("INSERT INTO accounts (acc_num, name, balance, active) VALUES (?,?,?,?)",
                            (acc_num, name, balance, active))
        self.connection.commit()

    def update_account(self, acc_num, name, balance):
        result = self.cursor.execute("SELECT balance FROM accounts WHERE acc_num = ?", (acc_num,))
        current_balance = result.fetchone()[0]
        if current_balance is None:
            logger.info("Account not in database. Adding account: %s - %s", acc_num, name)
            self.add_account(acc_num, name, balance)
        elif current_balance != balance:
            logger.info("Updating balance: %s - %s, R %.2f -> R %.2f", acc_num, name, current_balance, balance)
            self.cursor.execute("UPDATE accounts SET balance = ? WHERE acc_num = ?",
                                (float(balance), acc_num,))
            self.connection.commit()
        else:
            logger.debug("Account %s - %s up to date", acc_num, name)

    def set_account_inactive(self, acc_num):
        result = self.cursor.execute("SELECT * FROM accounts WHERE acc_num = ?", (acc_num,))
        account = result.fetchone()
        logger.info("Setting account status to inactive: %s - %s, balance = %.2f", account[0], account[1], account[2])
        self.cursor.execute("UPDATE accounts SET balance = ?, active = ? WHERE acc_num = ?", (0, False,))
        self.connection.commit()

    def add_transaction(self, acc_num, date, description, reference, amount, user_id, month=None, year=None):
        try:
            logger.debug("Adding transaction to database")
            self.cursor.execute('''INSERT INTO transactions (
                                        acc_num, date, description, reference, amount, 
                                        user_id, contrib_month, contrib_year)
                                   VALUES (?,?,?,?,?,?,?,?)''',
                                (acc_num, date, description, reference, amount, user_id, month, year))
            self.connection.commit()
        except sqlite3.IntegrityError:
            logger.debug("Transaction already in database")
            return False
        return self.cursor.lastrowid

    def check_transaction(self, acc_num, date, description, reference, amount):
        result = self.cursor.execute('''SELECT rowid
                                        FROM transactions
                                        WHERE acc_num = ? AND
                                              date = ? AND
                                              description = ? AND
                                              reference = ? AND
                                              amount = ?''',
                                     (acc_num, date, description, reference, amount))
        if not result.fetchone():
            logger.debug("Transaction not in database")
            return False
        else:
            logger.debug("Transaction already in database")
            return True

    def close_db(self):
        logger.info("Closing database connection!")
        self.cursor.close()
        self.connection.close()
        self.cursor = None
        self.connection = None
