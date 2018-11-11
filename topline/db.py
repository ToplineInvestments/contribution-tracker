import sqlite3
import json
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class DB:
    def __init__(self, filename='topline.db', user_json='users.json'):
        self.filename = filename
        if Path(filename).is_file():
            logger.debug("Connecting to database %s", self.filename)
            self.connection = sqlite3.connect(self.filename)
            self.cursor = self.connection.cursor()
        else:
            logger.warning('Database file not found: %s. Creating database', Path(filename).absolute())
            if Path(user_json).is_file():
                self.connection = sqlite3.connect(self.filename)
                self.initialise_db(user_json)
            else:
                logger.error('User file not found: %s. Unable to create database', Path(user_json).absolute())
                self.connection = None

    def initialise_db(self, user_json):
        logger.info("Initialising database tables")
        self.cursor.execute('''CREATE TABLE users(
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
        self.cursor.execute('''CREATE TABLE accounts(
                                            acc_num INTEGER PRIMARY KEY,
                                            name TEXT NOT NULL,
                                            balance REAL)
                            ''')
        self.cursor.execute('''CREATE TABLE transactions(
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
        self.cursor.execute('''CREATE UNIQUE INDEX unique_transaction 
                                            ON transactions(acc_num, date, description, reference, amount)
                            ''')

        self.connection.commit()
        if user_json:
            self.initialise_users(user_json)

    @staticmethod
    def get_users_from_json(json_filename):
        logger.debug("Reading user json file %s", json_filename)
        with open(json_filename, 'r') as f:
            users = json.load(f)
        return users

    def initialise_users(self, filename):
        logger.debug("Initialising users table")
        users = self.get_users_from_json(filename)
        for key, value in users.items():
            if 'alt_id' in value:
                key2 = value['alt_id']
            else:
                key2 = None
            logger.debug("Adding user: name = %s, surname = %s, email = %s, username = %s",
                         value['firstName'], value['lastName'], value['email'], key)
            self.cursor.execute(
                "INSERT INTO users (name, surname, email, username, alt_username) VALUES (?,?,?,?,?)",
                (value['firstName'], value['lastName'], value['email'], key, key2))
        self.connection.commit()

    def get_usernames(self):
        logger.debug("Getting usernames from database")
        result = self.cursor.execute("SELECT id, username, alt_username FROM users ORDER BY id")
        user_list = result.fetchall()
        # return [[uid[0], uid[1], uid[2]] if uid[2] else [uid[0], uid[1]] for uid in user_list]
        return user_list

    def update_user(self, user_id, username, date, amount, transaction_id):
        result = self.cursor.execute("SELECT total FROM users WHERE id = ?", (user_id,))
        total = result.fetchone()[0] or 0
        share = None
        logger.info("Updating user contribution from %s: R %.2f received on %s, total: R %.2f -> R %.2f",
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
        result = self.cursor.execute("SELECT acc_num, name FROM accounts")
        accounts = result.fetchall()
        logger.debug("Fetched %s accounts", len(accounts))
        return accounts

    def add_account(self, acc_num, name, balance):
        logger.info("Adding account to database: %s - %s, balance = R %.2f", acc_num, name, float(balance))
        self.cursor.execute("INSERT INTO accounts (acc_num, name, balance) VALUES (?,?,?)",
                            (acc_num, name, balance))
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

    def remove_account(self, acc_num):
        result = self.cursor.execute("SELECT * FROM accounts WHERE acc_num = ?", (acc_num,))
        account = result.fetchone()
        logger.info("Removing account from database: %s - %s, balance = %.2f", account[0], account[1], account[2])
        self.cursor.execute("DELETE FROM accounts WHERE acc_num = ?", (acc_num,))
        self.connection.commit()
