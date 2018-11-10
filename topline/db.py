import sqlite3
import json
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class DB:
    def __init__(self, filename='topline.db', user_json='users.json'):
        self.filename = filename
        if Path(filename).is_file():
            self.connection = sqlite3.connect(self.filename)
        else:
            logger.warning('Database file not found: %s. Creating database', Path(filename).absolute())
            if Path(user_json).is_file():
                self.connection = sqlite3.connect(self.filename)
                self.initialise_db(user_json)
            else:
                logger.error('User file not found: %s. Unable to create database', Path(user_json).absolute())
                self.connection = None

    def initialise_db(self, user_json_filename):
        self.connection.cursor().execute('''CREATE TABLE users(
                                            id INTEGER PRIMARY KEY, name TEXT NOT NULL, 
                                            surname TEXT NOT NULL,
                                            email TEXT UNIQUE NOT NULL,
                                            username TEXT UNIQUE NOT NULL,
                                            alt_username TEXT,
                                            last_transaction_id INTEGER,
                                            total REAL,
                                            share REAL)
                                         ''')
        self.connection.cursor().execute('''CREATE TABLE accounts(
                                            acc_num INTEGER PRIMARY KEY,
                                            name TEXT NOT NULL,
                                            balance REAL)
                                        ''')
        self.connection.cursor().execute('''CREATE TABLE transactions(
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
        self.connection.cursor().execute('''CREATE UNIQUE INDEX unique_transaction 
                                            ON transactions(acc_num, date, description, reference, amount)
                                        ''')
        users = self.get_users_from_json(user_json_filename)
        for key, value in users.items():
            if 'alt_id' in value:
                key2 = value['alt_id']
            else:
                key2 = None
            self.connection.cursor().execute(
                "INSERT INTO users (name, surname, email, username, alt_username) VALUES (?,?,?,?,?)",
                (value['firstName'], value['lastName'], value['email'], key, key2))
        self.connection.commit()

    @staticmethod
    def get_users_from_json(json_filename):
        with open(json_filename, 'r') as f:
            users = json.load(f)
        return users

    def get_user_ids(self):
        result = self.connection.cursor().execute("SELECT username, alt_username FROM users")
        user_list = result.fetchall()
        return [[uid[0], uid[1]] if uid[1] else [uid[0]] for uid in user_list]
