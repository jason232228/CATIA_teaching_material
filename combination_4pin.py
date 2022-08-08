import win32com.client as win32
import os

# ------------------------------讀取檔案路徑
system_root = os.path.dirname(os.path.realpath(__file__))
mother_part = system_root + '\\part\\'

# ------------------------------建立組立檔
catapp = win32.Dispatch("CATIA.Application")
document = catapp.Documents
productdoc = document.Add("Product")
product = productdoc.Product
products = product.Products

# ------------------------------匯入body
filedir = mother_part + 'body.CATPart'
import_file = filedir,
list(import_file)
productsvarient = products.AddComponentsFromFiles(import_file, "All")

# ------------------------------匯入pin
for i in range(0,4):
    filedir = mother_part + 'pin.CATPart'
    import_file = filedir,
    list(import_file)
    productsvarient = products.AddComponentsFromFiles(import_file, "All")

# ------------------------------連結pin與body的中心軸及底面
for i in range(1,5):
    constraints = product.Connections("CATIAConstraints")
    locals()['ref' + str(-3 + i * 4)] = product.CreateReferenceFromName("Product1/body.1/!PartBody/central_axis_%s" %i)
    locals()['ref' + str(-2 + i * 4)] = product.CreateReferenceFromName("Product1/pin.%s/!PartBody/central_axis" %i)
    constraint = constraints.AddBiEltCst(2, locals()['ref' + str(-3 + i * 4)], locals()['ref' + str(-2 + i * 4)])

    locals()['ref' + str(-1 + i * 4)] = product.CreateReferenceFromName("Product1/body.1/!body/!xy plane")
    locals()['ref' + str(0 + i * 4)] = product.CreateReferenceFromName("Product1/pin.%s/!pin/!xy plane" %i)
    constraint = constraints.AddBiEltCst(1, locals()['ref' + str(-1 + i * 4)], locals()['ref' + str(0 + i * 4)])

# ------------------------------更新組立狀態
length = constraint.Dimension
length.value = 0
constraint.Orientation = 0
product.Update()