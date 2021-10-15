from QueryGenerator import QueryGenerator

class TileInfoWriter:
    def __init__ (self, excelparser, sql, is_zno, update_query, tile_id):
        self.queryGenerator = QueryGenerator()
        self.excelparser = excelparser
        self.update_query = update_query
        self.sql = sql
        self.is_zno= is_zno
        self.name = self.excelparser.tile_name
        self.short_descr = self.excelparser.short_descr
        self.desc = self.excelparser.descr.replace('""', '"')
        self.tile_sn_id = self.excelparser.code
        self.tile_id = tile_id.lower()

    def get_tile_info(self):
        if self.is_zno:
            type_id = '1c0bc159-150a-e111-a31b-00155d04c01d'
        else:
            type_id = '1b0bc159-150a-e111-a31b-00155d04c01d'

        self.update_info_query = self.queryGenerator.update_tile_info(
            name=self.name, 
            short_desc=self.short_descr, 
            type_id=type_id, 
            desc=self.desc, 
            tile_id=self.tile_id
        )
        try:
            self.sql.manual(self.update_info_query)
        except:
            print(f"can't execute {self.update_info_query}")


    def write_tile_info(self):
        text = f'Link:	http://dev.nornickel/0/Nui/ViewModule.aspx#CardModuleV2/IteInfrastructureBackupTilePage/edit/{self.tile_id}\n\n\n'
        text += f'Id: {self.tile_id}\n'
        text += f'IteServiceNowId: {self.tile_sn_id}\n'
        text += f'IteName: {self.name}\n\n'
        text += f'IteCardSchema:\n\tName: Ite<>TilePage\n\tTile: Страница редактирования: «{self.name}»\n\n'
        text += f'ShortDescription: {self.short_descr}\n\n'
        text += f'Description:\n{self.desc}\n\n'
        try:
            text += f'UpdateQuery:\n\t{self.update_info_query}\n\n'
        except:
            text += 'UpdateQuery:\n\tCan`t set categoryies\n\n'
        try:
            text += f'\t{self.update_query}\n\n\n'
        except:
            text += f'\tCan`t update\n\n'
        text += 'DescriptionBuildIfoQuery:\n\n'
        text += f"DECLARE @TileId varchar(36) ='{self.tile_id}';"
        text += "INSERT INTO IteTileDescriptionBuildInfo\n"
        text += "(CreatedById, ModifiedById, ProcessListeners, IteTileId, IteColumnName, IteColumnTitle, IteOrder) VALUES\n"
        text += "('410006E1-CA4E-4502-A9EC-E54D922D2C00', '410006E1-CA4E-4502-A9EC-E54D922D2C00', "
        text += "0, @TileId, 'IteCategoryId', N'Тип запроса', 0),\n"
        text += "('410006E1-CA4E-4502-A9EC-E54D922D2C00', '410006E1-CA4E-4502-A9EC-E54D922D2C00', "
        text += "0, @TileId, 'IteDescription', N'Описание', 1);\n\n" 
        text += f"select * from IteTileDescriptionBuildInfo where IteTileId = '{self.tile_id}'"
        with open('out/tile.txt', 'w', encoding='utf-8') as f:
            f.write(text)
