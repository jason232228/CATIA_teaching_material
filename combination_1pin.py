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
filedir = "C:\\Users\\PDAL\\Desktop\\python組立catia教學\part\\body.CATPart"
import_file = filedir,
list(import_file)
productsvarient = products.AddComponentsFromFiles(import_file, "All")

# ------------------------------匯入pin
filedir = "C:\\Users\\PDAL\\Desktop\\python組立catia教學\part\\pin.CATPart"
import_file = filedir,
list(import_file)
productsvarient = products.AddComponentsFromFiles(import_file, "All")

# ------------------------------連結pin與body的中心軸
constraints = product.Connections("CATIAConstraints")
ref1 = product.CreateReferenceFromName("Product1/body.1/!PartBody/central_axis_1")
ref2 = product.CreateReferenceFromName("Product1/pin.1/!PartBody/central_axis")
constraint = constraints.AddBiEltCst(2, ref1, ref2)

# ------------------------------連結pin與body的底面（設定之間距離為0）
ref3 = product.CreateReferenceFromName("Product1/body.1/!body/!xy plane")
ref4 = product.CreateReferenceFromName("Product1/pin.1/!pin/!xy plane")
constraint = constraints.AddBiEltCst(1, ref3, ref4)

# ------------------------------更新組立狀態
length = constraint.Dimension
length.value = 0
constraint.Orientation = 0
product.Update()