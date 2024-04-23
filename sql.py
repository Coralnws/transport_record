# coding=utf-8
import sqlite3

def open_database():
    conn = sqlite3.connect('database.db')
    print("database open")
    c = conn.cursor()
    
    c.execute('''CREATE TABLE IF NOT EXISTS Record
        (id            integer primary key AUTOINCREMENT,
        date  DATE,
        lorry        TEXT    NOT NULL,
        reference       TEXT,
        company      TEXT    NOT NULL,
        description     TEXT,
        quantity        TEXT not null,
        unit_price      NUMERIC(10,2)	DEFAULT 0	NOT NULL,
        amount          NUMERIC(10,2)	DEFAULT 0	NOT NULL,
        created_at DATE,
        foreign key (lorry) references Lorry(lorry_number),
        foreign key (company) references Company(name) 
        );''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS Company
        (id            integer primary key AUTOINCREMENT,
        name           TEXT    NOT NULL,
        address_1 TEXT,
        address_2 TEXT,
        address_3 TEXT,
        address_4 TEXT,
        address_5 TEXT,
        tel TEXT
        );''')

    c.execute('''CREATE TABLE IF NOT EXISTS Employee
        (id            integer primary key AUTOINCREMENT,
        name           TEXT    NOT NULL,
        ic          TEXT    NOT NULL,
        lorry        TEXT    NOT NULL
        );''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS Lorry
        (id            integer primary key AUTOINCREMENT,
        lorry_number           TEXT    NOT NULL
        );''')

    c.execute('''CREATE TABLE IF NOT EXISTS Salary
        (id            integer primary key AUTOINCREMENT,
        date            DATE,
        employee           TEXT    NOT NULL,
        amount          NUMERIC(10,2)	DEFAULT 0	NOT NULL
        );''')

    c.execute('''CREATE TABLE IF NOT EXISTS Expenses_Company
        (id            integer primary key AUTOINCREMENT,
        name           TEXT    NOT NULL
         );''')

    c.execute('''CREATE TABLE IF NOT EXISTS Invoice
        (id            integer primary key AUTOINCREMENT,
        no           TEXT    NOT NULL,
        company        TEXT    NOT NULL,
        month           TEXT    NOT NULL,
        year            TEXT    NOT NULL,
        amount          NUMERIC(10,2)	DEFAULT 0	NOT NULL
         );''')


    c.execute('''CREATE TABLE IF NOT EXISTS Expenses
        (id            integer primary key AUTOINCREMENT,
        description           TEXT    NOT NULL
        );''')

    c.execute('''CREATE TABLE IF NOT EXISTS Expenses_Record
        (id            integer primary key AUTOINCREMENT,
        date           DATE,
        lorry        TEXT    NOT NULL,
        reference  TEXT    NOT NULL,
        expenses      TEXT    NOT NULL,
        company        TEXT NOT NULL,
        amount      NUMERIC(10,2)	DEFAULT 0	NOT NULL,
        foreign key (lorry) references Lorry(lorry_number),
        foreign key (expenses) references Expenses(description)
        )
        ''')
    
    print ("database create sucessfully")
    conn.commit()
    conn.close()

    