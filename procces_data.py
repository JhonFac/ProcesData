import pandas as pd

archivo_excel = "online_retail_II.xlsx"
hoja_2009_2010 = pd.read_excel(archivo_excel, sheet_name="Year 2009-2010")
hoja_2010_2011 = pd.read_excel(archivo_excel, sheet_name="Year 2010-2011")


# Función para encontrar el país que más productos consume en una hoja dada
def pais_mas_consumo(hoja):
    consumo_por_pais = hoja.groupby("Country")["Quantity"].sum().reset_index()
    pais_max_consumo = consumo_por_pais.loc[consumo_por_pais["Quantity"].idxmax()]
    return pais_max_consumo["Country"], pais_max_consumo["Quantity"]


# Encontrar el país que más productos consume en cada hoja
pais_max_consumo_2009_2010, cantidad_max_2009_2010 = pais_mas_consumo(hoja_2009_2010)
pais_max_consumo_2010_2011, cantidad_max_2010_2011 = pais_mas_consumo(hoja_2010_2011)

print("En la hoja 'Year 2009-2010':")
print("País que más productos consume:", pais_max_consumo_2009_2010)
print("Cantidad total de productos consumidos:", cantidad_max_2009_2010)

print("\nEn la hoja 'Year 2010-2011':")
print("País que más productos consume:", pais_max_consumo_2010_2011)
print("Cantidad total de productos consumidos:", cantidad_max_2010_2011)


##### ##############################
print("\n\n")


# Concatenar los datos de ambas hojas en un solo DataFrame
datos_combinados = pd.concat([hoja_2009_2010, hoja_2010_2011])

# Limpiar los datos y convertir la columna 'Price' a tipo numérico
datos_combinados["Price"] = datos_combinados["Price"].str.replace(",", ".").astype(float)

# Calcular las ganancias totales por producto
datos_combinados["TotalRevenue"] = datos_combinados["Quantity"] * datos_combinados["Price"]

# Calcular la cantidad total vendida por producto
productos_mas_vendidos = (
    datos_combinados.groupby(["StockCode", "Description"])[["Quantity", "TotalRevenue"]].sum().reset_index()
)

# Ordenar los productos por cantidad de ventas en orden descendente
productos_mas_vendidos = productos_mas_vendidos.sort_values(by="Quantity", ascending=False)

# Mostrar los productos más vendidos
print("Productos más vendidos por cantidad:")
print(productos_mas_vendidos.head())

# Ordenar los productos por ganancias en orden descendente
productos_mas_ganancias = productos_mas_vendidos.sort_values(by="TotalRevenue", ascending=False)

# Mostrar los productos más vendidos por ganancias
print("\nProductos más vendidos por ganancias:")
print(productos_mas_ganancias.head())
