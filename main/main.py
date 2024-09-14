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
        
        info_pedido = {
            f"{pnl['A2'].value}": [
                """ Sobre o Pedido """ 
                pnl['A2'].value,  # ID
                pnl['J2'].value,  # Data Pagamento
                
                """ Produtos """
                pnl['L2'].value,  # Nome do Produto
                pnl['M2'].value,  # SKU do Produto
                pnl['N2'].value,  # Variacao Produto
                pnl['O2'].value,  # Preco Original Produto
                pnl['P2'].value,  # Preco Acordadeo do Produto
                pnl['Q2'].value,  # Quantidade do Produto
                pnl['AL2'].value,  #Taxa Envio Reverso
                pnl['AM2'].value,  # Taxa Transacao
                pnl['AN2'].value,  # Taxa Comissao
                pnl['AO2'].value,  #  Taxa Servico
                
                """ Sobre o Cliente """
                pnl['AR2'].value,  # Username
                pnl['AY2'].value,  # Cidade Cliente
                pnl['AZ2'].value  # Estado Cliente
            ]
            
        }
        
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
    