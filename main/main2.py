import openpyxl.workbook
import pandas as pd
import openpyxl


class Main2:
    def __init__(self) -> None:
        pass
    
    
    
    
    def obtendo_valores():
        
        def save_produtos(linha_inserir, nome_produto, sku_produto):
            workbook_new = openpyxl.Workbook("RELATORIO_VENDAS.xlsx")
            sheet = workbook_new.active
            sheet[f"C{linha_inserir}"] = f"{nome_produto}"
            sheet[f"B{linha_inserir}"] = f"{sku_produto}"
            custo = input("informe o vlaor do produto: ")
            sheet[f"D{linha_inserir}"] = custo
            workbook_new.save("RELATORIO_VENDAS.xlsx")
        
        caminho = "pay07-14.xlsx"
        caminho_all = "all19-08a18-09.xlsx"
        workbook = openpyxl.load_workbook(caminho)
        workbook_all = openpyxl.load_workbook(caminho_all)
    
        pnl = workbook['Sheet1']
        pnl_all = workbook_all['orders']
        
        ctt = 0
        cont = 19
        total_valor = 0
        linha_add_produto = 4
        print("start")
        while ctt != 2:                           
            if f"{pnl[f"C{cont}"].value}" != "Saque" and ctt == 1:
                ### Valores que não sao o saque
                print(pnl[f"D{cont}"].value, pnl[f"F{cont}"].value)
                total_valor += float(pnl[f"F{cont}"].value)
                
                linha_achada = 1
                for linha in pnl_all["A"]:
                    if linha.value == pnl[f"D{cont}"].value:
                        print(pnl_all[f"M{linha_achada}"].value)
                        # save_produtos(linha_add_produto, 
                        #                     pnl_all[f"M{linha_achada}"].value, 
                        #                     pnl_all[f"N{linha_achada}"].value)
                        # linha_add_produto += 1
                    linha_achada += 1



                
            elif f"{pnl[f"C{cont}"].value}" == "Saque" and ctt == 0:
                total_valor -= float(pnl[f"H{cont}"].value)   
                
            if f"{pnl[f"C{cont}"].value}" == "Saque":
                ctt+=1
            cont+=1
        # print(pnl["A"].value)
        print(f"TOTAL DE SAQUE: {float(total_valor):.2f}")
Main2.obtendo_valores()