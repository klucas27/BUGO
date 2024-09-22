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
        valor = 0
        
        """Produto = [
            0 - nome
            1 - variacao
            2 - sku
        ]

        """
        for produtos in sheet_produtos['C']:
            # print(produto[0])
            # print(produto[1])
            if produto[0] == produtos.value and produto[1] == sheet_produtos['D'][linha].value:
                valor = sheet_produtos['E'][linha].value
                return float(valor)
            
            linha += 1        

        valor = Main.save_produtos([produto[2], # sku
                                    produto[0],  # nome 
                                    produto[1]]) # variacao
        return float(valor)
        
    
    def save_produtos(produto):
        """SALVA PRODUTOS NA TABELA RELATORIO VENDAS

        Args:
            produto (_list_): Eecebe lista com informações do produto
        """
        workbook_new = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_produtos = workbook_new["PRODUTOS"]
        df = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="PRODUTOS")
        nome = produto[1]
        variacao = produto[2]
        sku = produto[0]
        custo = 0
        
        linha_inserir = sheet_produtos.max_row + 1

        if variacao == "" or variacao == "nan":
            variacao = "S/V"
            
        # if nome in df.iloc[:, 2].values and variacao in df.iloc[:, 3].values:
        #     print("Entrou na 1 condi")
        # elif nome in df.iloc[:, 2].values and variacao == "S/V":
        #     print("Entrou na 2 condi")
        # else:
        sheet_produtos[f"B{linha_inserir}"] = f"{sku}"   #  SKU
        sheet_produtos[f"C{linha_inserir}"] = f"{nome}"   # Nome
        sheet_produtos[f"D{linha_inserir}"] = f"{variacao}"   # variação
        os.system("cls")
        print(f"\n -> {nome} \n  --> {variacao} \n  ---->{sku}") 
        custo = input("\ninforme o valor do custo do produto: ")
        sheet_produtos[f"E{linha_inserir}"] = float(custo)
        workbook_new.save("RELATORIO_VENDAS_fim.xlsx")
        return float(custo)
        
        
    def save_pedidos(infos_pedidos, id_igual = False):
        workbook_pedidos = openpyxl.load_workbook("RELATORIO_VENDAS.xlsx")
        sheet_vendas = workbook_pedidos["VENDAS"]
        df_vendas = pd.read_excel("RELATORIO_VENDAS.xlsx", sheet_name="VENDAS")
        linha_inserir_vendas = sheet_vendas.max_row + 1
        # print(f"Pedido de ID: {infos_pedidos[0]} Salvando...")
        sheet_vendas[f"B{linha_inserir_vendas}"] = f"{infos_pedidos[0]}"
        sheet_vendas[f"C{linha_inserir_vendas}"] = f"{infos_pedidos[1]}"
        sheet_vendas[f"D{linha_inserir_vendas}"] = f"{infos_pedidos[2]}"
        sheet_vendas[f"E{linha_inserir_vendas}"] = f"{infos_pedidos[3]}"
        sheet_vendas[f"F{linha_inserir_vendas}"] = f"{infos_pedidos[4]}"
        sheet_vendas[f"G{linha_inserir_vendas}"] = f"{infos_pedidos[5]}"
        sheet_vendas[f"H{linha_inserir_vendas}"] = f"{infos_pedidos[6]}"
        sheet_vendas[f"I{linha_inserir_vendas}"] = f"{infos_pedidos[7]}"
        sheet_vendas[f"J{linha_inserir_vendas}"] = f"{infos_pedidos[8]}"
        sheet_vendas[f"K{linha_inserir_vendas}"] = f"{infos_pedidos[9]}"
        sheet_vendas[f"L{linha_inserir_vendas}"] = f"{infos_pedidos[10]}"
        sheet_vendas[f"M{linha_inserir_vendas}"] = f"{infos_pedidos[11]}"
        sheet_vendas[f"N{linha_inserir_vendas}"] = f"{infos_pedidos[12]}"
        sheet_vendas[f"O{linha_inserir_vendas}"] = f"{infos_pedidos[13]}"
        sheet_vendas[f"P{linha_inserir_vendas}"] = f"{infos_pedidos[14]}"
        sheet_vendas[f"Q{linha_inserir_vendas}"] = f"{infos_pedidos[15]}"
        
        
        
        
        workbook_pedidos.save("RELATORIO_VENDAS.xlsx")
      

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
            if start_saque == 1 and str(ids._5) == "Entrada" and str(ids._3).startswith("Renda do pedido"):
                id_sacados.append(ids._4)

        ## Obtenção das informações dos pedidos conforme por id
        print("Obtendo informações do pedido...")
        caminho_all_oders = "all.xlsx"
        
        df_all_oders = pd.read_excel(caminho_all_oders, sheet_name="orders")
        # print(id_sacados.count("240831QTPQHE7X"))
        infos_obtidas = []
        id_pedido_only = []
        for pedidos_all in df_all_oders.itertuples():
            if id_sacados.count(f"{pedidos_all._1}") > 0:
                # print(pedidos_all)
                id_pedido_only.append(str(pedidos_all._1))
                variacao = None
                if str(pedidos_all._15) == "nan":
                    variacao = "S/V"
                else:
                    variacao = str(pedidos_all._15)
                    
                custo_produto = round(float(Main.consulta_preco([pedidos_all._13, variacao, pedidos_all._14])), 2)
                infos_obtidas.append([
                    pedidos_all._1,  #  0 ID
                    pedidos_all._2,  #  1 Status
                    pedidos_all._9,  #  2 dia do envio
                    pedidos_all._11,  # 3 dia pedido confirmado
                    pedidos_all._13,  #  4 Nome do produto
                    pedidos_all._14,  #  5 sku
                    variacao,  #  6 variação
                    pedidos_all._16,  #  7 preço original
                    pedidos_all._17,  #  8 preço de venda
                    pedidos_all.Quantidade,  #  9 quantidade
                    pedidos_all._20,  #  10 Sbtotal
                    pedidos_all._28,  #  11 gasto cupom
                    pedidos_all._33,  #  12 desconto + po -
                    pedidos_all._36,  #  13 Valor Total
                    pedidos_all._39,  #  14 Taxa reversa
                    pedidos_all._40,  #  15 Taxa transação
                    pedidos_all._41,  # 16 Taxa comissao
                    pedidos_all._42,  #  17 Taxa de serviço
                    pedidos_all.UF,  #  18 Estado do cliente
                    pedidos_all._57,  #  19 data completado
                    float(custo_produto)*int(pedidos_all.Quantidade),  # 20 Custo produto *                    
                    ])
    
        ## obtendo valores de custo
        print("Ordenando os pedidos...")
        caminho_save_relatorio = "RELATORIO_VENDAS.xlsx"
        df_produtos_save = pd.read_excel(caminho_save_relatorio, sheet_name="PRODUTOS")
        new_infos = []
        for pedidos_ordenados in id_sacados:
            for pedidos_desordenados in infos_obtidas:
                #print(pedidos_ordenados)
                if str(pedidos_desordenados[0]) == str(pedidos_ordenados):
                    new_infos.append(pedidos_desordenados)
                
                
        print("Salvando os pedidos na tabela!(Pode Demorar Bastante)...")
        ctt = 0
        while ctt < len(new_infos):
            print("ID: ", str(new_infos[ctt][0]), " Salvando...")
            tt_prod = int(id_pedido_only.count(str(new_infos[ctt][0])))

            total_venda = 0
            custo_total = 0
            while tt_prod >= 1:
                total_venda += new_infos[ctt][10]
                custo_total += new_infos[ctt][20]
                
                if tt_prod == 1:
                    valor_recebido = total_venda - (
                        float(new_infos[ctt][14]) +
                        float(new_infos[ctt][15]) +
                        float(new_infos[ctt][16]) +
                        float(new_infos[ctt][11]) +
                        float(new_infos[ctt][17])
                    )
                    
                    lucro_final = valor_recebido - custo_total
                    porct_lucro = lucro_final*100/custo_total
                    
                    Main.save_pedidos([
                        new_infos[ctt][0], #0
                        new_infos[ctt][1], #1
                        new_infos[ctt][18],  #2
                        new_infos[ctt][2],  # 3
                        new_infos[ctt][3],  # 4
                        new_infos[ctt][4],  # 5
                        new_infos[ctt][5],  # 6
                        new_infos[ctt][6],  # 7
                        new_infos[ctt][7],  # 8
                        round(float(new_infos[ctt][8]), 2),  #9
                        new_infos[ctt][9],  # 10
                        round(float(total_venda), 2),  # 11
                        round(float(valor_recebido), 2),  #12
                        round(float(custo_total), 2),  #13
                        round(float(lucro_final), 2),  # 14
                        round(float(porct_lucro), 2),  # 15
                                               
                    ])
                else:
                    Main.save_pedidos([
                        "||", #0
                        "||", #1
                        new_infos[ctt][18],  #2
                        new_infos[ctt][2],  # 3
                        new_infos[ctt][3],  # 4
                        new_infos[ctt][4],  # 5
                        new_infos[ctt][5],  # 6
                        new_infos[ctt][6],  # 7
                        new_infos[ctt][7],  # 8
                        round(float(new_infos[ctt][8]), 2),  #9
                        new_infos[ctt][9],  # 10
                        "||",  # 11
                        "||",  #12
                        "||",  #13
                        "||",  # 14
                        "||",  # 15
                    ])  
                    ctt+=1
                tt_prod -= 1
            ctt += 1
                    
        
Main.start_main()
