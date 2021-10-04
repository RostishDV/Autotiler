class ToGuidConverter:

	def fromfile(self):

		with open('resources/source.txt', 'r') as f:
		    ids = f.read().splitlines()

		guids = []
		for item_id in ids:
			guid = item_id[0:8] + '-' + item_id[8:12] + '-' + item_id[12:16] + '-' + item_id[16:20] + '-' + item_id[20:]

			guids.append(guid)

		print(guids)


	def fromString(self, id):

		guid = id[0:8] + '-' + id[8:12] + '-' + id[12:16] + '-' + id[16:20] + '-' + id[20:]
		return guid

def main():
	while True:
		inputedId = input()
		converter = ToGuidConverter()
		if inputedId == 'ex':
			break
		print(converter.fromString(id=inputedId))

if __name__ == '__main__':
	main()