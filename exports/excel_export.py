import xlsxwriter
import pandas as pd
from datetime import datetime
import os

STATUS_COLORS = {
    "Conciliada": "#27AE60",
    "Divergência de Valor": "#E67E22",
    "Sem Entrada": "#E74C3C",
    "Despesa": "#3498DB",
    "Devolução": "#9B59B6",
    "Pendente": "#95A5A6",
}

def export_excel(results, output_path, stats):
    wb = xlsxwriter.Workbook(output_path)

    # Formats
    header_fmt = wb.add_format({
        'bold': True, 'bg_color': '#1A2332', 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10
    })
    title_fmt = wb.add_format({
        'bold': True, 'font_size': 14, 'font_color': '#1A2332', 'bottom': 2
    })
    money_fmt = wb.add_format({'num_format': 'R$ #,##0.00', 'align': 'right'})
    date_fmt = wb.add_format({'align': 'center'})
    center_fmt = wb.add_format({'align': 'center'})
    pct_fmt = wb.add_format({'num_format': '0.0%', 'align': 'center'})

    color_fmts = {}
    for status, color in STATUS_COLORS.items():
        color_fmts[status] = wb.add_format({
            'bg_color': color + '33', 'font_color': '#1A2332',
            'border': 1, 'align': 'center', 'font_size': 9
        })
    row_fmt = wb.add_format({'border': 1, 'font_size': 9})
    row_money_fmt = wb.add_format({'border': 1, 'num_format': 'R$ #,##0.00', 'font_size': 9})

    def write_table(ws, data_list, title, status_label):
        ws.set_column('A:A', 5)
        ws.set_column('B:B', 15)
        ws.set_column('C:C', 35)
        ws.set_column('D:D', 18)
        ws.set_column('E:F', 15)
        ws.set_column('G:G', 15)
        ws.set_column('H:H', 20)
        ws.set_column('I:I', 30)

        ws.merge_range('A1:I1', title, title_fmt)
        ws.write('A2', f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', wb.add_format({'italic': True, 'font_color': '#666'}))

        headers = ['#', 'Nº Nota', 'Razão Social', 'CNPJ', 'Dt. Emissão', 'Valor ForLions', 'Valor Estoque', 'Classificação', 'Observações']
        for col, h in enumerate(headers):
            ws.write(3, col, h, header_fmt)

        for row_i, item in enumerate(data_list):
            r = row_i + 4
            ws.write(r, 0, row_i + 1, center_fmt)
            ws.write(r, 1, item.get('numero_nota', ''), row_fmt)
            ws.write(r, 2, item.get('razao_social', ''), row_fmt)
            ws.write(r, 3, item.get('cnpj', ''), row_fmt)
            ws.write(r, 4, item.get('data_emissao', ''), date_fmt)
            ws.write(r, 5, item.get('valor_forlions', 0) or 0, row_money_fmt)
            vs = item.get('valor_estoque')
            ws.write(r, 6, vs if vs is not None else '-', row_money_fmt if vs is not None else row_fmt)
            cls = item.get('classificacao', status_label)
            fmt = color_fmts.get(item.get('status', status_label), row_fmt)
            ws.write(r, 7, cls, fmt)
            ws.write(r, 8, item.get('observacoes', ''), row_fmt)

        ws.autofilter(3, 0, 3 + len(data_list), len(headers) - 1)
        return ws

    # Aba Conciliadas
    ws_c = wb.add_worksheet('✓ Conciliadas')
    ws_c.set_tab_color('#27AE60')
    write_table(ws_c, results.get('conciliadas', []), '✓ Notas Conciliadas', 'Conciliada')

    # Aba Divergências
    all_div = results.get('divergencias', []) + results.get('sem_entrada', [])
    ws_d = wb.add_worksheet('⚠ Divergências')
    ws_d.set_tab_color('#E74C3C')
    write_table(ws_d, all_div, '⚠ Divergências e Notas sem Entrada', 'Divergência')

    # Aba Despesas
    ws_e = wb.add_worksheet('$ Despesas')
    ws_e.set_tab_color('#3498DB')
    write_table(ws_e, results.get('despesas', []), '$ Despesas e Contas Operacionais', 'Despesa')

    # Aba Devoluções
    ws_dv = wb.add_worksheet('↩ Devoluções')
    ws_dv.set_tab_color('#9B59B6')
    write_table(ws_dv, results.get('devolucoes', []), '↩ Devoluções', 'Devolução')

    # Aba Resumo
    ws_r = wb.add_worksheet('📊 Resumo')
    ws_r.set_tab_color('#1A2332')
    ws_r.set_column('A:A', 35)
    ws_r.set_column('B:B', 20)

    ws_r.merge_range('A1:B1', '📊 RESUMO EXECUTIVO DE AUDITORIA FISCAL', title_fmt)
    ws_r.write('A2', f'Período da auditoria: {datetime.now().strftime("%d/%m/%Y")}',
               wb.add_format({'italic': True, 'font_color': '#666'}))

    kv_fmt = wb.add_format({'border': 1, 'font_size': 10, 'bold': True, 'bg_color': '#F8F9FA'})
    val_fmt = wb.add_format({'border': 1, 'font_size': 10, 'align': 'right'})
    val_money = wb.add_format({'border': 1, 'num_format': 'R$ #,##0.00', 'align': 'right'})
    val_pct = wb.add_format({'border': 1, 'num_format': '0.0%', 'align': 'right'})

    rows_r = [
        ('Total de Notas ForLions', stats.get('total_forlions', 0), val_fmt),
        ('Total Entradas Estoque', stats.get('total_estoque', 0), val_fmt),
        ('Notas Conciliadas', stats.get('total_conciliadas', 0), val_fmt),
        ('Divergências / Sem Entrada', stats.get('total_divergencias', 0), val_fmt),
        ('Despesas', stats.get('total_despesas', 0), val_fmt),
        ('Devoluções', stats.get('total_devolucoes', 0), val_fmt),
        ('Valor Total das Notas', stats.get('valor_total_notas', 0), val_money),
        ('Valor Conciliado', stats.get('valor_conciliado', 0), val_money),
        ('Percentual de Conformidade', stats.get('conformidade', 0) / 100 if stats.get('conformidade') else 0, val_pct),
    ]

    for i, (label, value, fmt) in enumerate(rows_r):
        ws_r.write(i + 3, 0, label, kv_fmt)
        ws_r.write(i + 3, 1, value, fmt)

    wb.close()
    return output_path
