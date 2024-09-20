import openpyxl.workbook
import pandas as pd
import openpyxl
import os

# 240907C70PXA8X

class Main2:
    def __init__(self) -> None:
        pass
    
    def save_produtos(produto, id_pedido):
        workbook_new = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_produtos = workbook_new["PRODUTOS"]
        sheet_vendas = workbook_new["VENDAS"]
        df = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="PRODUTOS")
        df_vendas = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="VENDAS")
    
        linha_inserir = sheet_produtos.max_row + 1
        linha_inserir_vendas = sheet_vendas.max_row + 1    
        # print(produto[2])
        #print(sheet_produtos.max_row)
        if produto[2] == "":
                produto[2] = "S/V"
        if produto[2] in df.iloc[:, 3].values and produto[1] in df.iloc[:, 2].values:
            pass
        elif produto[1] in df.iloc[:, 2].values and produto[2] == "":
            pass
        else:
            sheet_produtos[f"B{linha_inserir}"] = f"{produto[0]}"   #  SKU
            sheet_produtos[f"C{linha_inserir}"] = f"{produto[1]}"   # Nome
            sheet_produtos[f"D{linha_inserir}"] = f"{produto[2]}"   # variação
            custo = input("informe o valor do custo do produto: ")
            sheet_produtos[f"E{linha_inserir}"] = custo
        
        # SAVE PRODUTOS
        workbook_new.save("RELATORIO_VENDAS.xlsx")
        
        
        if id_pedido in df_vendas.iloc[:, 1].values and produto[1] in df_vendas.iloc[:, 3].values and produto[2] in df_vendas.iloc[:, 4].values:
            pass
        elif id_pedido in df_vendas.iloc[:, 1].values and produto[1] in df_vendas.iloc[:, 3].values and produto[2] == "":
            pass
        else:
            sheet_vendas[f"B{linha_inserir_vendas}"] = f"{id_pedido}"   #  ID
            sheet_vendas[f"C{linha_inserir_vendas}"] = f"{produto[0]}"  #  SKU
            sheet_vendas[f"D{linha_inserir_vendas}"] = f"{produto[1]}"  #  Nome
            sheet_vendas[f"E{linha_inserir_vendas}"] = f"{produto[2]}"  #  Variação
            sheet_vendas[f"F{linha_inserir_vendas}"] = f"{produto[3]}"  #  Quantidade produto
            sheet_vendas[f"G{linha_inserir_vendas}"] = f"{produto[4]}"  #  Estado do cliente
            sheet_vendas[f"H{linha_inserir_vendas}"] = f"{produto[5]}"  #  DATA DO PEDIDO
            sheet_vendas[f"I{linha_inserir_vendas}"] = f"{float(produto[6]):.2f}"  #  VALOR DO PEDIDO
            
            valor_pago_plataforma = float(produto[6]) - (float(produto[7]) + float(produto[8]) + float(produto[9]) + float(produto[10]) + float(produto[11]))
            
            sheet_vendas[f"J{linha_inserir_vendas}"] = valor_pago_plataforma  #  VALOR PAGO A PLATAFORMA
            
            custo_produto = 0
            line_find = 0
            #print(sheet_vendas["C"][3].value)
            for linha in sheet_produtos['C']:
                #print(linha.value)                
                if linha.value == produto[1] and f"{sheet_produtos[f"D{line_find+1}"].value}" == f"{produto[2]}":
                    # if linha.value == "S/V"
                    # # print(linha.value)
                    custo_produto = float(f"{sheet_produtos[f"E{line_find+1}"].value}".replace(",", "."))
                line_find +=1
            
            sheet_vendas[f"K{linha_inserir_vendas}"] = round(custo_produto*float(produto[3]), 2) #  CUSTO TOTAL DO PRODUTO

            lucro_final = valor_pago_plataforma - (custo_produto*float(produto[3]))
            sheet_vendas[f"L{linha_inserir_vendas}"] = round(lucro_final, 2)  #  VALOR PAGO A PLATAFORMA

            porcent_lucro = ((float(lucro_final)*100)/custo_produto)
            
            sheet_vendas[f"M{linha_inserir_vendas}"] = f"{round(porcent_lucro, 2)}%"  #  VALOR % LUCRO

            
        #SAVE TUDO
        workbook_new.save("RELATORIO_VENDAS.xlsx")
        
        
        
        
            
    #     print(custo_produto)
    #     sheet_vendas[f"H{linha_inserir_vendas}"] = f"{custo_produto}"  # Custo Produto
    #    # sheet_vendas[f"I{linha_inserir_vendas}"] = f"{produto[0]}"  #Lucro final
    #     workbook_new.save("RELATORIO_VENDAS.xlsx")
    
    def save_pedidos(linha_inserir, infos_pedidos):
        pass
        
        
    def obtendo_valores():     
        
        caminho = "pay.xlsx"
        caminho_all = "all.xlsx"
        workbook = openpyxl.load_workbook(caminho)
        workbook_all = openpyxl.load_workbook(caminho_all)
    
        pnl = workbook['Sheet1']
        pnl_all = workbook_all['orders']
        
        ctt = 0
        cont = 19
        total_valor = 0
        print("start")
        while ctt != 2:                           
            if f"{pnl[f"C{cont}"].value}" != "Saque" and ctt == 1:
                ### Valores que não sao o saque
                # print(pnl[f"D{cont}"].value, pnl[f"F{cont}"].value)
                total_valor += float(pnl[f"F{cont}"].value)
                
                linha_achada = 1
                for linha in pnl_all["A"]:
                    if linha.value == pnl[f"D{cont}"].value:
                        print(f"-> PRODUTO: {pnl_all[f"M{linha_achada}"].value}\n  -> VARIAÇÃO: {pnl_all[f"O{linha_achada}"].value}\n  -> SKU: {pnl_all[f"N{linha_achada}"].value}")
                        Main2.save_produtos([pnl_all[f"N{linha_achada}"].value,  #  0 SKU PRODUTO
                                            pnl_all[f"M{linha_achada}"].value,   #  1 NOME PRODUTO
                                            pnl_all[f"O{linha_achada}"].value,   #  2 VARIAÇÃO PRODUTO
                                            pnl_all[f"R{linha_achada}"].value,   #  3 QUANTIDADE
                                            pnl_all[f"BA{linha_achada}"].value,  #  4 ESTADO DO CLIENTE
                                            pnl_all[f"K{linha_achada}"].value,   #  5 DATA E HORA QUE O PEDIDO FOI LIBERADO
                                            pnl_all[f"T{linha_achada}"].value,   #  6 SUBTOTAL 
                                            pnl_all[f"AB{linha_achada}"].value,  #  7 CUPOM VENDEDOR  
                                            pnl_all[f"AC{linha_achada}"].value,  #  8 COIN CASHBACK  
                                            pnl_all[f"AG{linha_achada}"].value,  #  9 DESCONTO + por -  
                                            pnl_all[f"AO{linha_achada}"].value,  #  10 TAXA DE COMISSAO  
                                            pnl_all[f"AP{linha_achada}"].value,  #  11 TAXA DE SERVIÇO
                                            ], pnl_all[f"A{linha_achada}"].value,)
                    linha_achada += 1
                
                  
            
                os.system('cls')
            elif f"{pnl[f"C{cont}"].value}" == "Saque" and ctt == 0:
                total_valor -= float(pnl[f"H{cont}"].value)   
                
            if f"{pnl[f"C{cont}"].value}" == "Saque":
                ctt+=1
            cont+=1
        # print(pnl["A"].value)
        print(f"TOTAL DE SAQUE: {float(total_valor):.2f}")
Main2.obtendo_valores()