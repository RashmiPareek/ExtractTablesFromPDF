import tabula
import pandas


file = "C:\\Users\\SPECTRE\\Downloads\\ab.pdf"
#tabula.convert_into(file,"C:\\Users\\SPECTRE\\Downloads\\29-July-2021.csv",True)

dfs = tabula.read_pdf(file, pages="all")
print(dfs)

dfs[1].to_excel("C:\\Users\\SPECTRE\\Downloads\\ab.xlsx")
#print(dfs)
#tabula.convert_into_by_batch("C:\\Users\\SPECTRE\\Downloads","C:\\Users\\SPECTRE\\Downloads\\29-July-2021.csv",java_options=None)
