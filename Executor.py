from QueryGenerator import QueryGenerator
from Sql import Sql


class Executor():
    def __init__(self):
        database = 'dev.nornickel'
        self.connect = Sql(database=database)
        self.generator = QueryGenerator()


    def print_select_query_rows(self, rows, query, table):
        print("with query:")
        print(query)
        if len(rows) > 0:
            for row in rows:
                print(f"ID = {row[0]} in table = {table}")
                if table == "IteTopCategories":
                    self.category_section = row[0]
                if table == "IteGroupCategory":
                    self.category_group = row[0]
                if table == "IteTile":
                    self.tile_id = row[0]
                if len(row) > 1:
                    print(f"IteName = {row[1]}")
                print("====================================================================")
        else:
            print("------------------------ID NOT FOUND--------------------------------")    
            print("====================================================================")


    def execute_seource_lines(self, lines):
        for line in lines:
            attr = line.split("; ")
            table = attr[0]
            service_now = attr[1]
            if table == "IteTile":
                query = self.generator.getFromIteTileByServiceNowId(serviceNowId=service_now)
            else:
                query = self.generator.getFromTableByServiceNowId(serviceNowId=service_now, table_name=table)
            rows = self.connect.manual_select(query)
            self.print_select_query_rows(rows=rows, query=query, table=table)

    
    def read_source_file(self):
        with open('resources/source.txt', 'r') as f:
            lines = f.read().splitlines()
        return lines

    def try_set_categories_for_tile(self):
        try:
            self.tile_id
            self.category_group
            self.category_section
            self.set_categories_for_tile()
        except AttributeError:
            print("Can`t Update tile categories because no param id")


    def set_categories_for_tile(self):
        query = self.generator.update_tile_categories(tile_id=self.tile_id, category_group_id=self.category_group, category_section_id=self.category_section)
        print(query)
        self.connect.manual(query)


    def print_other_query_rows(self, rows, query, table):
        print("with query:")
        print(query)
        if len(rows) > 0:
            for row in rows:
                print(f"ID = {row[0].lower()} in table = {table}")
                print(f"IteName = {row[1]}")
                print("====================================================================")
        else:
            print("------------------------ID NOT FOUND--------------------------------")    
            print("====================================================================")


    def select_other_rows(self):
        with open('resources/other.txt', 'r') as f:
            lines = f.read().splitlines()
        for line in lines:
            attr = line.split("; ")
            table = attr[0]
            service_now = attr[1]
            query = self.generator.get_other_id_and_name_by_service_now(service_now=service_now, table_name=table)
            rows = self.connect.manual_select(query)
            self.print_other_query_rows(rows=rows, query=query, table=table)

