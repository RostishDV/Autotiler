import pyodbc
from datetime import datetime
from QueryGenerator import QueryGenerator
from ConfigReader import ConfigReader

class Sql:
    def __init__(self):
        configReader = ConfigReader()
        connectConfig = configReader.get_connect_config()
        server = connectConfig['server']
        db = connectConfig['db']
        driver = '{SQL Server Native Client 11.0}'
        self.cnxn = pyodbc.connect(f"Driver={driver};"+
                                   f"Server={server};"+
                                   f"Database={db};"+
                                   f"Trusted_Connection=yes;")
        self.query = "-- {}\n\n-- Made in Python".format(datetime.now()
                                                         .strftime("%d/%m/%Y"))


    def manual_select(self, query):
        cursor = self.cnxn.cursor()
        try:
            cursor.execute(query)
        except :
            print(f'exception in query = {query}')
        row = cursor.fetchone() 
        rows = []
        while row: 
            rows.append(row)
            row = cursor.fetchone()
        return rows


    def manual(self, query):
        cursor = self.cnxn.cursor()
        try:
            cursor.execute(query)
            self.cnxn.commit()
        except:
            print(f'exception in query = {query}')


def main():
    connection = Sql(database='dev.nornickel')
    serviceNowId = '73623a3fd90e6300bc5ba23f6290d343'
    generator = QueryGenerator()
    query = generator.getFromTableByServiceNowId(serviceNowId)
    response = connection.manual_select(query)
    print(response)

if __name__ == '__main__':
    main()
