from ToGuidConverter import ToGuidConverter

class QueryGenerator:

	def __init__(self):
		self.converter = ToGuidConverter()

	def getFromTableByServiceNowId(self, serviceNowId, table_name):
		service_now = self.converter.fromString(serviceNowId)
		query = "SELECT Id FROM "+ table_name +" WHERE IteServiceNowId='" + service_now + "'"
		return query

	def getFromIteTileByServiceNowId(self, serviceNowId):
		service_now = self.converter.fromString(serviceNowId)
		query = "SELECT Id, IteName FROM IteTile WHERE IteServiceNowId='" + service_now + "'"
		return query


	def update_tile_categories(self, tile_id, category_group_id, category_section_id):
		query = "UPDATE IteTile SET IteGroupCategoryId='"+ category_group_id + "'"
		query += ", IteCategorySectionId='" + category_section_id + "'"
		query += " WHERE Id='" + tile_id + "'"
		return query

	def get_other_id_and_name_by_service_now(self, service_now, table_name):
		service_now_id = self.converter.fromString(service_now)
		query = "SELECT Id, Name FROM "+ table_name +" WHERE IteServiceNowId='" + service_now_id + "'"
		return query


def main():
	serviceNowId = input()
	generator = QueryGenerator()
	print(generator.getIdFromIteTileByServiceNowId(serviceNowId))


if __name__ == '__main__':
 	main() 