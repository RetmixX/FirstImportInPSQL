import psycopg2
import openpyxl
import functionInsert as insert

def main():
    insert.InsertProductTypes()
    insert.InsertMaterialTypes()
    insert.InsertMaterials()
    insert.InsertProducts()
    insert.InsertProductMaterials()
    insert.book.close()

main()