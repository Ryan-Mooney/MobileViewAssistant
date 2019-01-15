import sqlite3, datetime

#Main function used to create table if needed and return connection to database
def connect():
    database=("assetDB.db")
    sql_create_assetData_table = """ CREATE TABLE IF NOT EXISTS AssetData (
                                        id integer PRIMARY KEY,
                                        asset NOT NULL,
                                        type text,
                                        pm_month_number integer,
                                        pm_month text,
                                        trial int,
                                        location text
                                    ); """

    sql_create_trial_table = """ CREATE TABLE IF NOT EXISTS TrialNumber (
                                        id integer PRIMARY KEY,
                                        trial_type text,
                                        date_ran text,
                                        time_ran text,
                                        date_time_ran text
                                    ); """

    connection=create_connection(database)
    if connection is not None:
        create_table(connection, sql_create_assetData_table)
        create_table(connection, sql_create_trial_table)
    else:
        print("Error! Could not create database connection!")
    return(connection)

#Takes each asset, determine what data descriptors are present, and saves them to the database
def save_to_db(assetList):
    for asset in assetList.keys():
        #Creates sql entry query for each asset based on the info provided
        asset_data=[asset]
        sql_text='asset'
        values_text='?'
        if assetList[asset]['Type']:
            asset_data.append(assetList[asset]['Type'])
            sql_text=sql_text+', type'
            values_text=values_text+', ?'
        if assetList[asset]['PM Month Number']:
            asset_data.append(assetList[asset]['PM Month Number'])
            sql_text=sql_text+', pm_month_number'
            values_text=values_text+', ?'
        if assetList[asset]['PM Month']:
            asset_data.append(assetList[asset]['PM Month'])
            sql_text=sql_text+', pm_month'
            values_text=values_text+', ?'
        asset_data.append(assetList[asset]['Trial'])
        sql_text=sql_text+', trial'
        values_text=values_text+', ?'
        asset_data.append(assetList[asset]['Location'])
        sql_text=sql_text+', location'
        values_text=values_text+', ?'
        #Saves each datapoint to database
        create_asset_data(asset_data, sql_text, values_text)

#Creates a unique trial number to identify each datapoint as being from the same trial and adds it to the assets dictionary data
def assign_trial_number(assetList, connection, trial_type):
    trial=create_trial_data(connection, trial_type)
    for asset in assetList:
        assetList[asset]['Trial']=trial
    return(assetList, trial)

#Creates a random datapoint for use in db testing
def create_random_datapoint():
    connection=create_connection("assetDB.db")
    asset_datapoint=('1000010', str(datetime.date.today()), str(datetime.datetime.now().time()), str(datetime.datetime.now()), 'Alaris IV Pump Module', '12', 'December', 'Floor 9')
    asset_id=create_asset_data(asset_datapoint)
    return(asset_id)

#Used to save datapoint to database and return its id if needed
def create_asset_data(asset_data, sql_text, values_text):
    connection=create_connection("assetDB.db")
    sql=' INSERT INTO AssetData('+sql_text+') VALUES('+values_text+') '
    cur=connection.cursor()
    cur.execute(sql, asset_data)
    return cur.lastrowid

#Creates and returns a unique trial number identifier
def create_trial_data(connection, trial_type):
    trial_type=[trial_type, str(datetime.date.today()), str(datetime.datetime.now().time()), str(datetime.datetime.now())]
    sql=''' INSERT INTO TrialNumber(trial_type, date_ran, time_ran, date_time_ran)
              VALUES(?, ?, ?, ?) '''
    cur=connection.cursor()
    cur.execute(sql, trial_type)
    return cur.lastrowid

#Returns all asset data from the database
def select_all_data():
    connection=create_connection("assetDB.db")
    cur=connection.cursor()
    cur.execute("SELECT * FROM AssetData")
    rows=cur.fetchall()
    for row in rows:
        print(row)
        
    cur.execute("SELECT * FROM TrialNumber")
    rows=cur.fetchall()
    for row in rows:
        print(row)

#Creates a database table
def create_table(connection, create_table_sql):
    try:
        c=connection.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)

#Connects to database
def create_connection(db_file):
    try:
        connection=sqlite3.connect(db_file, isolation_level=None)
        return(connection)
    except Error as e:
        print(e)
        return(None)
