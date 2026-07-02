from config import get_storage


storage = get_storage()

print(storage.listdir("./dados/xml_materiais_quirius"))
print(storage.listdir("//10.0.10.7/PUBLICO/vinicius.oliveira"))