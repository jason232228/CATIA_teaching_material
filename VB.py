# test_2：匯入body檔
#
# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# Set products1 = product1.Products
#
# Dim arrayOfVariantOfBSTR1(0)
# arrayOfVariantOfBSTR1(0) = "C:\Users\PDAL\Desktop\python組立catia教學\part\body.CATPart"
# products1.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"
#
# End Sub
#
# --------------------------------------------------
# test_3：匯入pin檔案
#
# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# Set products1 = product1.Products
#
# Dim arrayOfVariantOfBSTR1(0)
# arrayOfVariantOfBSTR1(0) = "C:\Users\PDAL\Desktop\python組立catia教學\part\pin.CATPart"
# products1.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"
#
# End Sub
#
# --------------------------------------------------
# test_4：同軸
#
# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# Set constraints1 = product1.Connections("CATIAConstraints")
#
# Set reference1 = product1.CreateReferenceFromName("Product1/body.1/!PartBody/central_axis_RT")
#
# Set reference2 = product1.CreateReferenceFromName("Product1/pin.1/!PartBody/central_axis")
#
# Set constraint1 = constraints1.AddBiEltCst(catCstTypeOn, reference1, reference2)
#
# End Sub
#
# --------------------------------------------------
# test_5：同面（面到面距離0）
#
# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# Set product1 = product1.ReferenceProduct
#
# Set constraints1 = product1.Connections("CATIAConstraints")
#
# Set reference1 = product1.CreateReferenceFromName("Product1/body.1/!body/xy plane")
#
# Set reference2 = product1.CreateReferenceFromName("Product1/pin.1/!pin/xy plane")
#
# Set constraint1 = constraints1.AddBiEltCst(catCstTypeDistance, reference1, reference2)
#
# Set length1 = constraint1.Dimension
#
# length1.Value = 0.000000
#
# constraint1.Orientation = catCstOrientSame
#
# End Sub
#
# --------------------------------------------------
# test_6：重整畫面
#
# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# product1.Update
#
# End Sub

# --------------------------------------------------
# test_ALL：完整阻立

# Language="VBSCRIPT"
#
# Sub CATMain()
#
# Set productDocument1 = CATIA.ActiveDocument
#
# Set product1 = productDocument1.Product
#
# Set products1 = product1.Products
#
# Dim arrayOfVariantOfBSTR1(0)
# arrayOfVariantOfBSTR1(0) = "C:\Users\PDAL\Desktop\python組立catia教學\part\body.CATPart"
# products1.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"
#
# Dim arrayOfVariantOfBSTR2(0)
# arrayOfVariantOfBSTR2(0) = "C:\Users\PDAL\Desktop\python組立catia教學\part\pin.CATPart"
# products1.AddComponentsFromFiles arrayOfVariantOfBSTR2, "All"
#
# Set constraints1 = product1.Connections("CATIAConstraints")
#
# Set reference1 = product1.CreateReferenceFromName("Product1/body.1/!PartBody/central_axis_RT")
#
# Set reference2 = product1.CreateReferenceFromName("Product1/pin.1/!PartBody/central_axis")
#
# Set constraint1 = constraints1.AddBiEltCst(catCstTypeOn, reference1, reference2)
#
# Set constraints1 = product1.Connections("CATIAConstraints")
#
# Set reference3 = product1.CreateReferenceFromName("Product1/body.1/!body/xy plane")
#
# Set reference4 = product1.CreateReferenceFromName("Product1/pin.1/!pin/xy plane")
#
# Set constraint2 = constraints1.AddBiEltCst(catCstTypeDistance, reference3, reference4)
#
# Set length1 = constraint2.Dimension
#
# length1.Value = 0.000000
#
# constraint2.Orientation = catCstOrientSame
#
# product1.Update
#
# End Sub