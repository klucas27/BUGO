import pandas as pd
import openpyxl
print("Software para analises de dados!\n")


class Bugo():
    def __init__(self) -> None:
        escolha = input("Informe a opção: ")
        if escolha == 1:
            self.get_infos()
        pass
    
    def get_infos():
        
        # caminho = input("Informe o caminho: ")
        caminho = "C:\\Users\\PC FEIRAO 03\\Desktop\\Order.shipping.20240912_20240912.xlsx"
        workbook = openpyxl.load_workbook(caminho)
        pnl = workbook['orders']
        
        # print(pnl.max_row)
        ctt = 2
        while ctt <= pnl.max_row:
            info_pedido = {
                f"{pnl[f'A{ctt}'].value}": [
                    # Sobre o Pedido  
                    pnl[f'A{ctt}'].value,  # ID
                    pnl[f'J{ctt}'].value,  # Data Pagamento
                    
                    # Produtos 
                    pnl[f'L{ctt}'].value,  # Nome do Produto
                    pnl[f'M{ctt}'].value,  # SKU do Produto
                    pnl[f'N{ctt}'].value,  # Variacao Produto
                    pnl[f'O{ctt}'].value,  # Preco Original Produto
                    pnl[f'P{ctt}'].value,  # Preco Acordadeo do Produto
                    pnl[f'Q{ctt}'].value,  # Quantidade do Produto
                    pnl[f'AL{ctt}'].value,  #Taxa Envio Reverso
                    pnl[f'AM{ctt}'].value,  # Taxa Transacao
                    pnl[f'AN{ctt}'].value,  # Taxa Comissao
                    pnl[f'AO{ctt}'].value,  #  Taxa Servico
                    
                    # Sobre o Cliente
                    pnl[f'AR{ctt}'].value,  # Username
                    pnl[f'AY{ctt}'].value,  # Cidade Cliente
                    pnl[f'AZ{ctt}'].value,  # Estado Cliente
                ]
            }
            ctt += 1
            for pas in info_pedido.values():
                print(pas[0])
                
                # for pas2 in pas:
                #     print(pas2[1])
            
        
        # """ Sobre o Pedido """       
        # id_pedido = pnl['A2'].value
        # data_pagamento = pnl['J2'].value
        
        # """ Produtos """
        # produto_nome = pnl['L2'].value
        # produto_sku_ref = pnl['M2'].value
        # produto_variacao = pnl['N2'].value
        # produto_preco_orig = pnl['O2'].value
        # produto_preco_acordado = pnl['P2'].value
        # produto_quantidade = pnl['Q2'].value
        # produto_taxa_envio_reverso = pnl['AL2'].value
        # produto_taxa_transacao = pnl['AM2'].value
        # produto_taxa_comicao = pnl['AN2'].value
        # produto_taxa_servico = pnl['AO2'].value
        
        # """ Sobre o Cliente """
        # cli_username = pnl['AR2'].value
        # cli_cidade = pnl['AY2'].value
        # cli_estado = pnl['AZ2'].value
        
        
        
if __name__ == "__main__":
    Bugo.get_infos()
    