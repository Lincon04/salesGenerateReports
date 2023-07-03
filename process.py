import locale

import pandas as pd
from rows_convert import ColumnsConvert
import time
from datetime import datetime

locale.setlocale(locale.LC_ALL, '')


class Vendas:
    def __init__(self, file_name):
        self.df_vendas = pd.read_json(file_name)
        self.pv = 'pv'
        self.data_veda = 'data venda'
        self.nsu = 'nsu'
        self.rv_number = 'rv number'
        self.valor_bruto = 'valor bruto'
        self.desconto = 'desconto'
        self.valor_liquido = 'valor liquido'
        self.modalidade = 'modalidade'
        self.meio_de_pagamento = 'meio de pagamento'
        self.maquininha = 'maquininha'
        self.quantidade_parcelas = 'quantidade parcela(s)'
        self.parcela = 'parcela(s)'
        self.data_recebimento = 'data recebimento'

    # def rename_columns(self):
    #     self.df_vendas.columns = [self.pv, self.data_veda, self.nsu, self.rv_number, self.valor_bruto, self.desconto,
    #                               self.valor_liquido, self.modalidade, self.meio_de_pagamento, self.maquininha,
    #                               self.quantidade_parcelas, self.parcela, self.data_recebimento]

    def convert_date(self, field):
        self.df_vendas[field] = pd.to_datetime(self.df_vendas[field])

    def format_float_number(self, field):
        field_series = self.df_vendas[field]
        self.df_vendas[field] = field_series.apply(lambda x: locale.currency(x, grouping=True))

    def convertendo_para_numeric(self, field):
        self.df_vendas[field] = pd.to_numeric(field)

    def __str__(self):
        return f'{self.df_vendas.head(10)}'

    def process_data(self):
        # self.rename_columns()

        for x in ['data_venda', 'data_recebimento']:
            self.convert_date(x)

        # for x in ['valor_bruto', 'desconto', 'valor_liquido']:
        #     self.format_float_number(x)
        # self.convertendo_para_numeric(x)

        print(type(self.df_vendas['valor_bruto'][1]))

    def save_data(self):
        self.df_vendas.reset_index(drop=True, inplace=True)
        writer = pd.ExcelWriter('../reports/excel_vendas.xlsx', engine='openpyxl')
        self.df_vendas.to_excel(writer, index=False,
                                header=[self.pv, self.data_veda, self.nsu, self.rv_number, self.valor_bruto,
                                        self.desconto, self.valor_liquido, self.modalidade, self.meio_de_pagamento,
                                        self.maquininha, self.quantidade_parcelas, self.parcela, self.data_recebimento],
                                sheet_name='Vendas')

        convert = ColumnsConvert(writer, 'Vendas')
        convert.add_format_currency([5, 6, 7])
        convert.add_width_columns()
        convert.create_style_to_header()
        convert.salvar('../reports/excel_vendas2.xlsx')


vendas = Vendas('../temp/Relatorio-Vendas-js-82703.json')
vendas.process_data()
vendas.save_data()
