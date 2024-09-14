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
        
        """ Sobre o Pedido """       
        id_pedido = pnl['A2'].value
        data_pagamento = pnl['J2'].value
        
        """ Produtos """
        produto_nome = pnl['L2'].value
        produto_sku_ref = pnl['M2'].value
        produto_variacao = pnl['N2'].value
        produto_preco_orig = pnl['O2'].value
        produto_preco_acordado = pnl['P2'].value
        produto_quantidade = pnl['Q2'].value
        produto_taxa_envio_reverso = pnl['AL2'].value
        produto_taxa_transacao = pnl['AM2'].value
        produto_taxa_comicao = pnl['AN2'].value
        produto_taxa_servico = pnl['AO2'].value
        
        """ Sobre o Cliente """
        cli_username = pnl['AR2'].value
        cli_cidade = pnl['AY2'].value
        cli_estado = pnl['AZ2'].value
        
        
        
if __name__ == "__main__":
    Bugo.get_infos()
    