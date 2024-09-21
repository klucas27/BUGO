import openpyxl.workbook
import pandas as pd
import openpyxl
import os

# 240907C70PXA8X

class Main:
    def __init__(self) -> None:
        pass
   
    def consulta_preco(produto):
        workbook_new = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_produtos = workbook_new["PRODUTOS"]
        df = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="PRODUTOS")
        
        linha = 0
        for produtos in sheet_produtos['C']:
            # print(produto[0])
            # print(produto[1])
            if produto[0] == produtos.value and produto[1] == sheet_produtos['D'][linha].value:
                valor = sheet_produtos['E'][linha].value
                return float(valor)
            linha += 1
        
        return 0
        
    
    def save_produtos(produto):
        """SALVA PRODUTOS NA TABELA RELATORIO VENDAS

        Args:
            produto (_list_): Eecebe lista com informações do produto
        """
        workbook_new = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_produtos = workbook_new["PRODUTOS"]
        # sheet_vendas = workbook_new["VENDAS"]
        df = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="PRODUTOS")
        # df_vendas = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="VENDAS")
    
        linha_inserir = sheet_produtos.max_row + 1
        # linha_inserir_vendas = sheet_vendas.max_row + 1    
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
        
        
        # if id_pedido in df_vendas.iloc[:, 1].values and produto[1] in df_vendas.iloc[:, 3].values and produto[2] in df_vendas.iloc[:, 4].values:
        #     pass
        # elif id_pedido in df_vendas.iloc[:, 1].values and produto[1] in df_vendas.iloc[:, 3].values and produto[2] == "":
        #     pass
        # else:
        #     sheet_vendas[f"B{linha_inserir_vendas}"] = f"{id_pedido}"   #  ID
        #     sheet_vendas[f"C{linha_inserir_vendas}"] = f"{produto[0]}"  #  SKU
        #     sheet_vendas[f"D{linha_inserir_vendas}"] = f"{produto[1]}"  #  Nome
        #     sheet_vendas[f"E{linha_inserir_vendas}"] = f"{produto[2]}"  #  Variação
        #     sheet_vendas[f"F{linha_inserir_vendas}"] = f"{produto[3]}"  #  Quantidade produto
        #     sheet_vendas[f"G{linha_inserir_vendas}"] = f"{produto[4]}"  #  Estado do cliente
        #     sheet_vendas[f"H{linha_inserir_vendas}"] = f"{produto[5]}"  #  DATA DO PEDIDO
        #     sheet_vendas[f"I{linha_inserir_vendas}"] = f"{float(produto[6]):.2f}"  #  VALOR DO PEDIDO
            
        #     valor_pago_plataforma = float(produto[6]) - (float(produto[7]) + float(produto[8]) + float(produto[9]) + float(produto[10]) + float(produto[11]))
            
        #     sheet_vendas[f"J{linha_inserir_vendas}"] = valor_pago_plataforma  #  VALOR PAGO A PLATAFORMA
            
        #     custo_produto = 0
        #     line_find = 0
        #     #print(sheet_vendas["C"][3].value)
        #     for linha in sheet_produtos['C']:
        #         #print(linha.value)                
        #         if linha.value == produto[1] and f"{sheet_produtos[f"D{line_find+1}"].value}" == f"{produto[2]}":
        #             # if linha.value == "S/V"
        #             # # print(linha.value)
        #             custo_produto = float(f"{sheet_produtos[f"E{line_find+1}"].value}".replace(",", "."))
        #             print(custo_produto)
        #         line_find +=1
            
        #     sheet_vendas[f"K{linha_inserir_vendas}"] = round(custo_produto*float(produto[3]), 2) #  CUSTO TOTAL DO PRODUTO

        #     lucro_final = valor_pago_plataforma - (custo_produto*float(produto[3]))
        #     sheet_vendas[f"L{linha_inserir_vendas}"] = round(lucro_final, 2)  #  VALOR PAGO A PLATAFORMA

        #     try:
        #         porcent_lucro = ((float(lucro_final)*100)/custo_produto)
        #         sheet_vendas[f"M{linha_inserir_vendas}"] = f"{round(porcent_lucro, 2)}%"  #  VALOR % LUCRO                
        #     except:
        #         os.system(f"Echo 'ERROr: Lucro: {lucro_final} | Custo: {custo_produto}'")
        #         os.system("Pause")
                

            
        # #SAVE TUDO
        # workbook_new.save("RELATORIO_VENDAS.xlsx")
        
        
        
        
            
    #     print(custo_produto)
    #     sheet_vendas[f"H{linha_inserir_vendas}"] = f"{custo_produto}"  # Custo Produto
    #    # sheet_vendas[f"I{linha_inserir_vendas}"] = f"{produto[0]}"  #Lucro final
    #     workbook_new.save("RELATORIO_VENDAS.xlsx")
    
    def save_pedidos(infos_pedidos, id_igual = False):
        workbook_pedidos = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_vendas = workbook_pedidos["VENDAS"]
        df_vendas = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="VENDAS")
        linha_inserir_vendas = sheet_vendas.max_row + 1
        
        if id_igual == False:
            if infos_pedidos[0] in df_vendas.iloc[:, 1].values and infos_pedidos[2] in df_vendas.iloc[:, 3].values and infos_pedidos[3] in df_vendas.iloc[:, 4].values:
                pass
            elif infos_pedidos[0] in df_vendas.iloc[:, 1].values and infos_pedidos[2] in df_vendas.iloc[:, 3].values and infos_pedidos[3] == "S/V":
                pass
            else:
                sheet_vendas[f"B{linha_inserir_vendas}"] = f"{infos_pedidos[0]}"   #  ID Pedido
                sheet_vendas[f"C{linha_inserir_vendas}"] = f"{infos_pedidos[1]}"   #  SKU
                sheet_vendas[f"D{linha_inserir_vendas}"] = f"{infos_pedidos[2]}"   #  PRODUTO
                sheet_vendas[f"E{linha_inserir_vendas}"] = f"{infos_pedidos[3]}"   #  VARIAÇÃO
                sheet_vendas[f"F{linha_inserir_vendas}"] = f"{infos_pedidos[4]}"   #  QUANTIDADE
                sheet_vendas[f"G{linha_inserir_vendas}"] = f"{infos_pedidos[5]}"   #  ESTADO COMPRADOR
                sheet_vendas[f"H{linha_inserir_vendas}"] = f"{infos_pedidos[6]}"   #  DATA E HORA
                sheet_vendas[f"I{linha_inserir_vendas}"] = f"{round(float(infos_pedidos[7]), 2)}"   #  VALOR DO PRODUTO
                sheet_vendas[f"J{linha_inserir_vendas}"] = f"{round(float(infos_pedidos[8]), 2)}"   #  VALOR PEDIDO
                sheet_vendas[f"K{linha_inserir_vendas}"] = f"{round(float(infos_pedidos[9]), 2)}"   #  VALOR PAGO DA PLATAFORMA
                sheet_vendas[f"L{linha_inserir_vendas}"] = f"{round(float(infos_pedidos[10]), 2)}"   #  Custo produto
                
                # custo_produto = round((Main.consulta_preco([infos_pedidos[2], infos_pedidos[3]])*int(infos_pedidos[4])), 2)
                # print(f"Custo do produto: {(Main.consulta_preco([infos_pedidos[2], infos_pedidos[3]]))}")
                # sheet_vendas[f"L{linha_inserir_vendas}"] = f"{custo_produto}"   #  CUSTO DO PRODUTO
                
                lucro_final = round(infos_pedidos[9] - infos_pedidos[10], 2)
                sheet_vendas[f"M{linha_inserir_vendas}"] = f"{lucro_final}"   #  LUCRO FINAL
                perc_lucro = 0
                if infos_pedidos[10] == 0:
                    perc_lucro = 0
                else:
                    perc_lucro = round((lucro_final*100)/infos_pedidos[10], 2)
                
                sheet_vendas[f"N{linha_inserir_vendas}"] = f"{perc_lucro}%"   #  % DE LUCRO
                workbook_pedidos.save("RELATORIO_VENDAS.xlsx")
        else:
                sheet_vendas[f"B{linha_inserir_vendas}"] = f"{infos_pedidos[0]}"   #  ID Pedido
                sheet_vendas[f"C{linha_inserir_vendas}"] = f"{infos_pedidos[1]}"   #  SKU
                sheet_vendas[f"D{linha_inserir_vendas}"] = f"{infos_pedidos[2]}"   #  PRODUTO
                sheet_vendas[f"E{linha_inserir_vendas}"] = f"{infos_pedidos[3]}"   #  VARIAÇÃO
                sheet_vendas[f"F{linha_inserir_vendas}"] = f"{infos_pedidos[4]}"   #  QUANTIDADE
                sheet_vendas[f"G{linha_inserir_vendas}"] = f"{infos_pedidos[5]}"   #  ESTADO COMPRADOR
                sheet_vendas[f"H{linha_inserir_vendas}"] = f"{infos_pedidos[6]}"   #  DATA E HORA
                sheet_vendas[f"I{linha_inserir_vendas}"] = f"{infos_pedidos[7]}"   #  VALOR DO PRODUTO
                sheet_vendas[f"J{linha_inserir_vendas}"] = f"{infos_pedidos[8]}"   #  VALOR PEDIDO
                sheet_vendas[f"K{linha_inserir_vendas}"] = f"{infos_pedidos[9]}"   #  VALOR PAGO DA PLATAFORMA
                sheet_vendas[f"L{linha_inserir_vendas}"] = f"{infos_pedidos[10]}"   #  custo Produto
                
                # custo_produto = round((Main.consulta_preco([infos_pedidos[2], infos_pedidos[3]])*int(infos_pedidos[4])), 2)
                # sheet_vendas[f"L{linha_inserir_vendas}"] = f"{custo_produto}"   #  CUSTO DO PRODUTO
                workbook_pedidos.save("RELATORIO_VENDAS.xlsx")
                
                # lucro_final = infos_pedidos[9] - custo_produto
                # sheet_vendas[f"M{linha_inserir_vendas}"] = f"{lucro_final}"   #  LUCRO FINAL
                
                # perc_lucro = (lucro_final*100)/custo_produto
                
                # sheet_vendas[f"N{linha_inserir_vendas}"] = f"{perc_lucro}%"  
    
        # workbook_pedidos.save("RELATORIO_VENDAS.xlsx")
               

    def start_main():
        caminho_pay = "pay.xlsx"
        df_pay = pd.read_excel(caminho_pay, sheet_name="Sheet1")
                
        ## Obtenção dos ids Sacados
        start_saque = 0
        id_sacados = []
        print("Obtendo ids Sacados...")        
        for ids in df_pay.itertuples():
            if ids._3 == "Saque":
                start_saque +=1
                continue
            if start_saque == 1:
                id_sacados.append(ids._4)

        ## Obtenção das informações dos pedidos conforme por id
        print("Obtendo informações do pedido...")
        caminho_all_oders = "all.xlsx"
        
        df_all_oders = pd.read_excel(caminho_all_oders, sheet_name="orders")
        # print(id_sacados.count("240831QTPQHE7X"))
        infos_obtidas = []
        for pedidos_all in df_all_oders.itertuples():
            if id_sacados.count(f"{pedidos_all._1}") > 0:
                print(pedidos_all)
                infos_obtidas.append([
                    pedidos_all._1,  #  ID
                    pedidos_all._2,  # Status
                    pedidos_all._9,  # dia do envio
                    pedidos_all._11,  # dia pedido confirmado
                    pedidos_all._13,  # Nome do produto
                    pedidos_all._14,  # sku
                    pedidos_all._15,  # variação
                    pedidos_all._16,  # preço original
                    pedidos_all._17,  # preço de venda
                    pedidos_all.Quantidade,  # quantidade
                    pedidos_all._20,  # Sbtotal
                    pedidos_all._28,  # gasto cupom
                    pedidos_all._33,  # desconto + po -
                    pedidos_all._36,  # Valor Total
                    pedidos_all._39,  # Taxa reversa
                    pedidos_all._40,  # Taxa transação
                    pedidos_all._41,  # Taxa comissao
                    pedidos_all._42,  # Taxa de serviço
                    pedidos_all.UF,  # Estado do cliente
                    pedidos_all._57,  # data completado                    
                    ])
        
        print(infos_obtidas)
    
        ## obtendo valores de custo
        caminho_save_relatorio = "RELATORIO_VENDAS.xlsx"
        df_produtos_save = pd.read_excel(caminho_save_relatorio, sheet_name="PRODUTOS")
        
        
        
        
    def obtendo_valores():     
        
        caminho_pay = "pay.xlsx"
        caminho_all = "all.xlsx"
        caminho_relat_vendas = "RELATORIO_VENDAS.xlsx"
        
        workbook_pay = openpyxl.load_workbook(caminho_pay)
        workbook_all = openpyxl.load_workbook(caminho_all)
        
        df_oders = pd.read_excel("all.xlsx", sheet_name="orders")
        df_vendas = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="VENDAS")
        df_produtos = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="PRODUTOS")
        df_pay = pd.read_excel("pay.xlsx", sheet_name="Sheet1")
                
        tp_oders = df_oders.itertuples()
        tp_pay = list(df_pay.itertuples())                
        start = 0
        
        for pas in tp_pay:
            print("\n", pas.Pandas[2])
        
        
        # for ids_pay in tp_pay:
        #     if str(ids_pay._3) == "Saque":
        #         start += 1
        #     if start == 1:
        #         for ids_oders in df_oders.itertuples():
        #             if ids_pay._4 == ids_oders._1:
        #                 if df_produtos.iloc[:, 3].isin([""])
                        
                        
        #                 print(ids_oders)
                                    
        #     if start > 1:
        #         break
            
        
        
        # print("----MENU----")
        # escolha_menu = 0
        # print("1 - Para Obter os produtos")
        # print("2 - Para Obter os pedidos")
        # escolha_menu = input(">> ")
        # print(escolha_menu)
        # if int(escolha_menu) == 1:
        #     print("Entrou na opção 1!")
        #     ctt = 0
        #     cont = 19
        #     id_obtido = None          
            
        #     while ctt != 2:                           
        #         if f"{pnl[f"C{cont}"].value}" != "Saque" and ctt == 1: 
                                    
        #             linha_achada = 1
                    
        #             for linha in pnl_all["A"]: # pnl_all da tabela all

        #                 if linha.value == pnl[f"D{cont}"].value:
                            
        #                     if (df_vendas.iloc[:, 1] == linha.value).sum() > 0:
        #                         continue
        #                     else:
                            
        #                         print(f"-> PRODUTO: {pnl_all[f"M{linha_achada}"].value}\n  -> VARIAÇÃO: {pnl_all[f"O{linha_achada}"].value}\n  -> SKU: {pnl_all[f"N{linha_achada}"].value}")
                                
        #                         Main.save_produtos([pnl_all[f"N{linha_achada}"].value, #  0 SKU PRODUTO
        #                                             pnl_all[f"M{linha_achada}"].value, #  1 NOME PRODUTO
        #                                             pnl_all[f"O{linha_achada}"].value  #  2 VARIAÇÃO PRODUTO
        #                                             ])
        #                         variacao = None
        #                         if pnl_all[f"O{linha_achada}"].value == "":
        #                             variacao = "S/V"
        #                         else:
        #                             variacao = pnl_all[f"O{linha_achada}"].value
                                    
        #                         total_de_vezes = (df_oders.iloc[:, 0] == pnl_all[f"A{linha_achada}"].value).sum()
        #                         # print(total_de_vezes)
        #                         valor_total_pedido = 0
        #                         custo_total_pedido = 0
        #                         if total_de_vezes == 1:
        #                             valor_total_pedido = float(pnl_all[f"T{linha_achada}"].value)
        #                             custo_total_pedido = round((Main.consulta_preco([pnl_all[f"M{linha_achada}"].value, variacao])*int(pnl_all[f"R{linha_achada}"].value)), 2)                                 
        #                             Main.save_pedidos([
        #                                 pnl_all[f"A{linha_achada}"].value,  #  0 ID PEDIDO
        #                                 pnl_all[f"N{linha_achada}"].value,  #  1 SKU PRODUTO
        #                                 pnl_all[f"M{linha_achada}"].value,  #  2 NOME DO PRODUTO
        #                                 variacao,  #  3 VARIAÇÃO PRODUTO
        #                                 pnl_all[f"R{linha_achada}"].value,  #  4 QUANTIDADE DE PRODUTO
        #                                 pnl_all[f"BA{linha_achada}"].value, #  5 ESTADO DO COMPRADOR
        #                                 pnl_all[f"BE{linha_achada}"].value, #  6 DATA E HORAS
        #                                 float(pnl_all[f"Q{linha_achada}"].value), # 7 VALOR VENDA DO PRODUTO
        #                                 valor_total_pedido, #  8 VALOR TOTAL DO PEDIDO
        #                                 (float(pnl_all[f"AJ{linha_achada}"].value) - 
        #                                 (float(pnl_all[f"AO{linha_achada}"].value) + 
        #                                 float(pnl_all[f"AP{linha_achada}"].value)+
        #                                 float(pnl_all[f"AB{linha_achada}"].value)+
        #                                 float(pnl_all[f"AG{linha_achada}"].value))), #  9 VALOR TOTAL PAGO A PLATAFORMA
        #                                 float(custo_total_pedido)  # 10 Custo de produtos
        #                             ])
                                
        #                         else:
        #                             for pas in range(0, total_de_vezes):
        #                                 valor_total_pedido += float(pnl_all[f"T{linha_achada+pas}"].value)
        #                                 custo_total_pedido += float(round((Main.consulta_preco([pnl_all[f"M{linha_achada+pas}"].value, variacao])*int(pnl_all[f"R{linha_achada+pas}"].value)), 2))
        #                                 # os.system("Pause")
                                    
        #                             Main.save_pedidos([
        #                                 pnl_all[f"A{linha_achada}"].value,  #  0 ID PEDIDO
        #                                 pnl_all[f"N{linha_achada}"].value,  #  1 SKU PRODUTO
        #                                 pnl_all[f"M{linha_achada}"].value,  #  2 NOME DO PRODUTO
        #                                 variacao,  #  3 VARIAÇÃO PRODUTO
        #                                 pnl_all[f"R{linha_achada}"].value,  #  4 QUANTIDADE DE PRODUTO
        #                                 pnl_all[f"BA{linha_achada}"].value, #  5 ESTADO DO COMPRADOR
        #                                 pnl_all[f"BE{linha_achada}"].value, #  6 DATA E HORAS
        #                                 float(pnl_all[f"Q{linha_achada}"].value), # 7 VALOR VENDA DO PRODUTO
        #                                 valor_total_pedido, #  8 VALOR TOTAL DO PEDIDO
        #                                 (float(valor_total_pedido) - (float(pnl_all[f"AO{linha_achada}"].value) + 
        #                                 float(pnl_all[f"AP{linha_achada}"].value)+
        #                                 float(pnl_all[f"AB{linha_achada}"].value)+
        #                                 float(pnl_all[f"AG{linha_achada}"].value))), #  9 VALOR TOTAL PAGO A PLATAFORMA
        #                                 float(custo_total_pedido)  # 10 Custo total de mercadoria
        #                             ])
        #                             linha_achada += 1
        #                             for pas in range(1, total_de_vezes):
        #                                 Main.save_pedidos([
        #                                 "||",  #  0 ID PEDIDO
        #                                 pnl_all[f"N{linha_achada}"].value,  #  1 SKU PRODUTO
        #                                 pnl_all[f"M{linha_achada}"].value,  #  2 NOME DO PRODUTO
        #                                 variacao,  #  3 VARIAÇÃO PRODUTO
        #                                 pnl_all[f"R{linha_achada}"].value,  #  4 QUANTIDADE DE PRODUTO
        #                                 "||", #  5 ESTADO DO COMPRADOR
        #                                 "||", #  6 DATA E HORAS
        #                                 float(pnl_all[f"Q{linha_achada}"].value), # 7 VALOR VENDA DO PRODUTO
        #                                 "||", #  8 VALOR TOTAL DO PEDIDO
        #                                 "||", #  9 VALOR TOTAL PAGO A PLATAFORMA
        #                                 "||"  # 10 Custo Mercadoria
        #                                 ], id_igual=True)
        #                                 linha_achada += 1
                                    
                            
                                        
                                
        #                     # id_obtido = pnl_all[f"A{linha_achada}"].value
        #                 linha_achada += 1
        #             # os.system('cls')
        #         # elif f"{pnl[f"C{cont}"].value}" == "Saque" and ctt == 0:
        #         #     total_valor -= float(pnl[f"H{cont}"].value)   
                    
        #         if f"{pnl[f"C{cont}"].value}" == "Saque":
        #             ctt+=1
        #         cont+=1
        #     # print(pnl["A"].value)
        #     # print(f"TOTAL DE SAQUE: {float(total_valor):.2f}")
        # elif escolha_menu == 2:
        #     Main.save_pedidos()
        
        
Main.start_main()

# produto = "Kit Chave Torx + Kit Allen 18 Peças Aço Cromo Vanadium"
# variacao = "Allen + Torx"

# valor = Main.consulta_preco([produto, variacao])

# print("Retornado:  -->>", valor)