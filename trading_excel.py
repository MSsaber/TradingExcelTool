#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   trading_exce;.py
@Time    :   2023/02/12 19:18:44
@Author  :   xiaobai
@Version :   1.0
@Contact :   1752615737@qq.com
@Desc    :   解析表格
'''

import copy
import datetime
try:
    import xlrd
    import xlwt
except Exception as e:
    print(e)

deal_num_col = 6
deal_price_col = 7
deal_amount_col = 8

_excel_style_dir = {}

def init_default_excel_style():
    global _excel_style_dir
    #sum table style
    sumtable_header_style = xlwt.XFStyle()  # Create the pattern
    sumtable_header_font = xlwt.Font()
    sumtable_header_font.name = u"微软黑雅"
    sumtable_header_font.bold = True
    sumtable_header_font.height = 20*18
    sumtable_header_style.font = sumtable_header_font
    sumtable_header_alignment = xlwt.Alignment()
    sumtable_header_alignment.horz = 0x02
    sumtable_header_alignment.vert = 0x01
    sumtable_header_style.alignment = sumtable_header_alignment
    _excel_style_dir['sumtable_tital'] = sumtable_header_style
    
    sumtable_date_style = xlwt.XFStyle()  # Create the pattern
    sumtable_date_font = xlwt.Font()
    sumtable_date_font.name = u"等线"
    sumtable_date_font.bold = True
    sumtable_date_font.height = 20*14
    sumtable_date_style.font = sumtable_date_font
    sumtable_date_alignment = xlwt.Alignment()
    sumtable_date_alignment.horz = 0x02
    sumtable_date_alignment.vert = 0x01
    sumtable_date_style.alignment = sumtable_date_alignment
    _excel_style_dir['sumtable_date'] = sumtable_date_style

    sumtable_style = xlwt.XFStyle()  # Create the pattern
    sumtable_pattern = xlwt.Pattern()  # Create the pattern
    # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    sumtable_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    sumtable_pattern.pattern_fore_colour = 0x25
    sumtable_font = xlwt.Font()
    sumtable_font.name = u"宋体"
    sumtable_font.bold = True
    sumtable_font.height = 20*12
    sumtable_style.font = sumtable_font
    sumtable_alignment = xlwt.Alignment()
    sumtable_alignment.horz = 0x03
    sumtable_alignment.vert = 0x01
    sumtable_style.alignment = sumtable_alignment
    sumtable_style.pattern = sumtable_pattern  # Add pattern to style
    _excel_style_dir['sumtable_zero'] = sumtable_style

    sumtable_theader_style = xlwt.XFStyle()  # Create the pattern
    sumtable_theader_font = xlwt.Font()
    sumtable_theader_font.name = u"宋体"
    sumtable_theader_font.bold = True
    sumtable_theader_font.height = 20*12
    sumtable_theader_style.font = sumtable_theader_font
    sumtable_theader_borders = xlwt.Borders()
    sumtable_theader_borders.left = xlwt.Borders.THIN
    sumtable_theader_borders.right = xlwt.Borders.THIN
    sumtable_theader_borders.top = xlwt.Borders.THIN
    sumtable_theader_borders.bottom = xlwt.Borders.THIN
    sumtable_theader_borders.left_colour = 0xff
    sumtable_theader_borders.right_colour = 0xff
    sumtable_theader_borders.top_colour = 0xff
    sumtable_theader_borders.bottom_colour = 0xff
    sumtable_theader_style.borders = sumtable_theader_borders
    sumtable_theader_alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    sumtable_theader_alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    sumtable_theader_alignment.vert = 0x01
    sumtable_theader_style.alignment = sumtable_theader_alignment
    _excel_style_dir['sumtable_header'] = sumtable_theader_style

    sumtable_cell_style = xlwt.XFStyle()
    sumtable_cell_borders = xlwt.Borders()
    sumtable_cell_borders.left = xlwt.Borders.THIN
    sumtable_cell_borders.right = xlwt.Borders.THIN
    sumtable_cell_borders.top = xlwt.Borders.THIN
    sumtable_cell_borders.bottom = xlwt.Borders.THIN
    sumtable_cell_borders.left_colour = 0xff
    sumtable_cell_borders.right_colour = 0xff
    sumtable_cell_borders.top_colour = 0xff
    sumtable_cell_borders.bottom_colour = 0xff
    sumtable_cell_style.borders = sumtable_cell_borders
    sumtable_cell_font = xlwt.Font()
    sumtable_cell_font.name = u"宋体"
    sumtable_cell_font.height = 20*12
    sumtable_cell_style.font = sumtable_cell_font
    sumtable_cell_alignment = xlwt.Alignment()
    sumtable_cell_alignment.horz = 0x03
    sumtable_cell_alignment.vert = 0x01
    sumtable_cell_style.alignment = sumtable_cell_alignment
    _excel_style_dir['sumtable_cell'] = sumtable_cell_style

    sumtable_cell_buy_style = xlwt.XFStyle()
    sumtable_cell_buy_borders = xlwt.Borders()
    sumtable_cell_buy_borders.left = xlwt.Borders.THIN
    sumtable_cell_buy_borders.right = xlwt.Borders.THIN
    sumtable_cell_buy_borders.top = xlwt.Borders.THIN
    sumtable_cell_buy_borders.bottom = xlwt.Borders.THIN
    sumtable_cell_buy_borders.left_colour = 0xff
    sumtable_cell_buy_borders.right_colour = 0xff
    sumtable_cell_buy_borders.top_colour = 0xff
    sumtable_cell_buy_borders.bottom_colour = 0xff
    sumtable_cell_buy_style.borders = sumtable_cell_buy_borders
    sumtable_cell_buy_font = xlwt.Font()
    sumtable_cell_buy_font.name = u"宋体"
    sumtable_cell_buy_font.height = 20*12
    sumtable_cell_buy_font.colour_index = 0x27
    sumtable_cell_buy_style.font = sumtable_cell_buy_font
    sumtable_cell_buy_alignment = xlwt.Alignment()
    sumtable_cell_buy_alignment.horz = 0x03
    sumtable_cell_buy_alignment.vert = 0x01
    sumtable_cell_buy_style.alignment = sumtable_cell_buy_alignment
    _excel_style_dir['sumtable_cell_buy'] = sumtable_cell_buy_style

    sumtable_cell_sell_style = xlwt.XFStyle()
    sumtable_cell_sell_borders = xlwt.Borders()
    sumtable_cell_sell_borders.left = xlwt.Borders.THIN
    sumtable_cell_sell_borders.right = xlwt.Borders.THIN
    sumtable_cell_sell_borders.top = xlwt.Borders.THIN
    sumtable_cell_sell_borders.bottom = xlwt.Borders.THIN
    sumtable_cell_sell_borders.left_colour = 0xff
    sumtable_cell_sell_borders.right_colour = 0xff
    sumtable_cell_sell_borders.top_colour = 0xff
    sumtable_cell_sell_borders.bottom_colour = 0xff
    sumtable_cell_sell_style.borders = sumtable_cell_sell_borders
    sumtable_cell_sell_font = xlwt.Font()
    sumtable_cell_sell_font.name = u"宋体"
    sumtable_cell_sell_font.height = 20*12
    sumtable_cell_sell_font.colour_index = 0x26
    sumtable_cell_sell_style.font = sumtable_cell_sell_font
    sumtable_cell_sell_alignment = xlwt.Alignment()
    sumtable_cell_sell_alignment.horz = 0x03
    sumtable_cell_sell_alignment.vert = 0x01
    sumtable_cell_sell_style.alignment = sumtable_cell_sell_alignment
    _excel_style_dir['sumtable_cell_sell'] = sumtable_cell_sell_style

    sumtable_total_style = xlwt.XFStyle()
    sumtable_total_pattern = xlwt.Pattern()  # Create the pattern
    sumtable_total_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    sumtable_total_pattern.pattern_fore_colour = 0x28
    sumtable_total_style.pattern = sumtable_total_pattern
    sumtable_total_borders = xlwt.Borders()
    sumtable_total_borders.left = xlwt.Borders.THIN
    sumtable_total_borders.right = xlwt.Borders.THIN
    sumtable_total_borders.top = xlwt.Borders.THIN
    sumtable_total_borders.bottom = xlwt.Borders.THIN
    sumtable_total_borders.left_colour = 0xff
    sumtable_total_borders.right_colour = 0xff
    sumtable_total_borders.top_colour = 0xff
    sumtable_total_borders.bottom_colour = 0xff
    sumtable_total_style.borders = sumtable_total_borders
    sumtable_total_font = xlwt.Font()
    sumtable_total_font.name = u"宋体"
    sumtable_total_font.height = 20*12
    sumtable_total_style.font = sumtable_total_font
    sumtable_total_alignment = xlwt.Alignment()
    sumtable_total_alignment.horz = 0x03
    sumtable_total_alignment.vert = 0x01
    sumtable_total_style.alignment = sumtable_total_alignment
    _excel_style_dir['sumtable_total'] = sumtable_total_style

    #total table
    totaltable_zero_buy_style = xlwt.XFStyle()  # Create the pattern
    totaltable_zero_buy_pattern = xlwt.Pattern()  # Create the pattern
    totaltable_zero_buy_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    totaltable_zero_buy_pattern.pattern_fore_colour = 0x21
    totaltable_zero_buy_font = xlwt.Font()
    totaltable_zero_buy_font.name = u"等线"
    totaltable_zero_buy_font.bold = True
    totaltable_zero_buy_font.height = 20*14
    totaltable_zero_buy_style.font = totaltable_zero_buy_font
    totaltable_zero_buy_alignment = xlwt.Alignment()
    totaltable_zero_buy_alignment.horz = 0x03
    totaltable_zero_buy_alignment.vert = 0x01
    totaltable_zero_buy_style.alignment = totaltable_zero_buy_alignment
    totaltable_zero_buy_style.pattern = totaltable_zero_buy_pattern  # Add pattern to style
    _excel_style_dir['totaltable_zero_buy'] = totaltable_zero_buy_style

    totaltable_zero_sell_style = xlwt.XFStyle()  # Create the pattern
    totaltable_zero_sell_pattern = xlwt.Pattern()  # Create the pattern
    totaltable_zero_sell_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    totaltable_zero_sell_pattern.pattern_fore_colour = 0x22
    totaltable_zero_sell_font = xlwt.Font()
    totaltable_zero_sell_font.name = u"等线"
    totaltable_zero_sell_font.bold = True
    totaltable_zero_sell_font.height = 20*14
    totaltable_zero_sell_style.font = totaltable_zero_sell_font
    totaltable_zero_sell_alignment = xlwt.Alignment()
    totaltable_zero_sell_alignment.horz = 0x03
    totaltable_zero_sell_alignment.vert = 0x01
    totaltable_zero_sell_style.alignment = totaltable_zero_sell_alignment
    totaltable_zero_sell_style.pattern = totaltable_zero_sell_pattern  # Add pattern to style
    _excel_style_dir['totaltable_zero_sell'] = totaltable_zero_sell_style

    totaltable_header_style = xlwt.XFStyle()  # Create the pattern
    totaltable_header_font = xlwt.Font()
    totaltable_header_font.name = u"等线"
    totaltable_header_font.bold = True
    totaltable_header_font.height = 20*12
    totaltable_header_style.font = totaltable_header_font
    totaltable_header_alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    totaltable_header_alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    totaltable_header_alignment.vert = 0x01
    totaltable_header_style.alignment = totaltable_header_alignment
    _excel_style_dir['totaltable_header'] = totaltable_header_style

    totaltable_data_style = xlwt.XFStyle()
    totaltable_data_font = xlwt.Font()
    totaltable_data_font.name = u"等线"
    totaltable_data_font.height = 20*11
    totaltable_data_style.font = totaltable_data_font
    totaltable_data_alignment = xlwt.Alignment()
    totaltable_data_alignment.horz = 0x02
    totaltable_data_alignment.vert = 0x01
    totaltable_data_style.alignment = totaltable_data_alignment
    _excel_style_dir['totaltable_data'] = totaltable_data_style

    #single strategy table
    strategy_zero_buy_style = xlwt.XFStyle()  # Create the pattern
    strategy_zero_buy_pattern = xlwt.Pattern()  # Create the pattern
    strategy_zero_buy_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    strategy_zero_buy_pattern.pattern_fore_colour = 0x21
    strategy_zero_buy_font = xlwt.Font()
    strategy_zero_buy_font.name = u"等线"
    strategy_zero_buy_font.bold = True
    strategy_zero_buy_font.height = 20*14
    strategy_zero_buy_style.font = strategy_zero_buy_font
    strategy_zero_buy_alignment = xlwt.Alignment()
    strategy_zero_buy_alignment.horz = 0x03
    strategy_zero_buy_alignment.vert = 0x01
    strategy_zero_buy_style.alignment = strategy_zero_buy_alignment
    strategy_zero_buy_style.pattern = strategy_zero_buy_pattern
    _excel_style_dir['strategy_zero_buy'] = strategy_zero_buy_style

    strategy_zero_sell_style = xlwt.XFStyle()  # Create the pattern
    strategy_zero_sell_pattern = xlwt.Pattern()  # Create the pattern
    strategy_zero_sell_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    strategy_zero_sell_pattern.pattern_fore_colour = 0x22
    strategy_zero_sell_font = xlwt.Font()
    strategy_zero_sell_font.name = u"等线"
    strategy_zero_sell_font.bold = True
    strategy_zero_sell_font.height = 20*14
    strategy_zero_sell_style.font = strategy_zero_sell_font
    strategy_zero_sell_alignment = xlwt.Alignment()
    strategy_zero_sell_alignment.horz = 0x03
    strategy_zero_sell_alignment.vert = 0x01
    strategy_zero_sell_style.alignment = strategy_zero_sell_alignment
    strategy_zero_sell_style.pattern = strategy_zero_sell_pattern
    _excel_style_dir['strategy_zero_sell'] = strategy_zero_sell_style

    strategy_header_style = xlwt.XFStyle()  # Create the pattern
    strategy_header_font = xlwt.Font()
    strategy_header_font.name = u"等线"
    strategy_header_font.bold = True
    strategy_header_font.height = 20*12
    strategy_header_style.font = strategy_header_font
    strategy_header_alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    strategy_header_alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    strategy_header_alignment.vert = 0x01
    strategy_header_style.alignment = strategy_header_alignment
    _excel_style_dir['strategy_header'] = strategy_header_style

    strategy_data_style = xlwt.XFStyle()
    strategy_data_font = xlwt.Font()
    strategy_data_font.name = u"等线"
    strategy_data_font.height = 20*11
    strategy_data_style.font = strategy_data_font
    strategy_data_alignment = xlwt.Alignment()
    strategy_data_alignment.horz = 0x02
    strategy_data_alignment.vert = 0x01
    strategy_data_style.alignment = strategy_data_alignment
    _excel_style_dir['strategy_data'] = strategy_data_style

    strategy_data_buy_style = xlwt.XFStyle()
    strategy_data_buy_font = xlwt.Font()
    strategy_data_buy_font.name = u"等线"
    strategy_data_buy_font.height = 20*11
    strategy_data_buy_font.colour_index = 0x26
    strategy_data_buy_style.font = strategy_data_buy_font
    strategy_data_buy_alignment = xlwt.Alignment()
    strategy_data_buy_alignment.horz = 0x02
    strategy_data_buy_alignment.vert = 0x01
    strategy_data_buy_style.alignment = strategy_data_buy_alignment
    _excel_style_dir['strategy_data_buy'] = strategy_data_buy_style

    strategy_data_sell_style = xlwt.XFStyle()
    strategy_data_sell_font = xlwt.Font()
    strategy_data_sell_font.name = u"等线"
    strategy_data_sell_font.height = 20*11
    strategy_data_sell_font.colour_index = 0x27
    strategy_data_sell_style.font = strategy_data_sell_font
    strategy_data_sell_alignment = xlwt.Alignment()
    strategy_data_sell_alignment.horz = 0x02
    strategy_data_sell_alignment.vert = 0x01
    strategy_data_sell_style.alignment = strategy_data_sell_alignment
    _excel_style_dir['strategy_data_sell'] = strategy_data_sell_style


def add_excel_style(key, style, cover = True):
    global _excel_style_dir
    if key in _excel_style_dir.keys():
        if cover is True:
            _excel_style_dir[key] = style
        else:
            return
    else:
        _excel_style_dir[key] = style

def get_excel_style(key):
    global _excel_style_dir
    return _excel_style_dir[key]

class TableData:
    def __init__(self, name, header=None, datas=None):
        self.name = name
        self.zerorow = []
        self.header = header
        self.datas = datas

    def import_data_by_sheet(self, sheet):
        self.datas = []
        needins = False
        for i in range(sheet.nrows):
            if i == 0:
                self.header = sheet.row_values(i)
                if "序号" not in self.header:
                    self.header = ["序号"] + self.header
                    needins = True
            else:
                datas = sheet.row_values(i)
                if needins:
                    datas = [str(i)] + datas
                self.datas.append(datas)

    def set_header(self, header):
        self.header = header

    def set_datas(self, datas):
        self.datas = datas

    def set_number(self):
        for i in range(len(self.datas)):
            self.datas[i] = [str(i+1)] + self.datas[i]

    def reset_number(self):
        for i in range(len(self.datas)):
            self.datas[i][0] = str(i+1)

    def fomat_col(self, col_lsit, t):
        for i in range(len(self.datas)):
            for c in col_lsit:
                v = self.datas[i][c]
                if t == float or t == int:
                    s = ''
                    if isinstance(v, str) and ',' in v:
                        vs = v.split(',')
                        for sub in vs:
                            s += sub
                        v = s
                self.datas[i][c] = t(v)

    def get_col_by_tital(self, titals):
        tital = titals.split(';')
        for i in range(len(self.header)):
            if self.header[i] in tital:
                return i
        return -1

    def merge_row_by_col(self, col_lsit, col_merge):
        keymap = {}
        for row in self.datas:
            key = ''
            for col in col_lsit:
                key += str(row[col])
            if key not in keymap.keys():
                keymap[key] = []
            keymap[key].append(row)
        res = []
        for k, v in keymap.items():
            new_row = []
            last = len(v) - 1
            for c in range(len(v[0])):
                if c not in col_merge:
                    new_row.append(v[last][c])
                else:
                    fill = None
                    for i in range(len(v)):
                        if i == 0:
                            fill = v[i][c]
                        else:
                            fill += v[i][c]
                    new_row.append(fill)
            res.append(new_row)
        self.datas = res

    def merge_row_by_tital(self, tital_list, col_merge):
        col_list = []
        for tital in tital_list:
            for i in range(len(self.header)):
                if self.header[i] == tital:
                    col_list.append(i)
        if len(col_list) > 0:
            self.merge_row_by_col(col_list, col_merge)
            return True
        return False

    def split_table_by_col(self, col):
        keymap = {}
        if self.datas is None:
            return None
        for row in self.datas:
            if row[col] not in keymap.keys():
                keymap[row[col]] = []
            keymap[row[col]].append(row)
        res = []
        for k, v in keymap.items():
            table = TableData(self.name + '-' + k[len(k)-9:])
            table.set_header(self.header)
            table.set_datas(v)
            res.append(table)
        return res

    def split_table_by_tital(self, titals):
        tital = titals.split(';')
        for i in range(len(self.header)):
            if self.header[i] in tital:
                return self.split_table_by_col(i)
        return None

    def show_datas(self):
        print(self.name)
        print(self.header)
        for row in self.datas:
            print(row)

    @staticmethod
    def gen_tabledata(name, sheet):
        table = TableData(name)
        table.import_data_by_sheet(sheet)
        return table

    def _merge_trading(self):
        remove_row = []
        dictcol = self.get_col_by_tital('方向')
        numcol = self.get_col_by_tital('成交数量')

        for row in range(len(self.datas)):
            if row in remove_row:
                continue
            if row != len(self.datas) - 1:
                flag = self.datas[row][dictcol]
                if flag != self.datas[row + 1][dictcol] and \
                        self.datas[row][numcol] + self.datas[row + 1][numcol] == 0:
                    remove_row.append(row)
                    remove_row.append(row + 1)
                    break
                elif flag != self.datas[row + 1][dictcol] and \
                        self.datas[row][numcol] + self.datas[row + 1][numcol] != 0 and \
                        row + 2 != len(self.datas) - 1:
                    removeEnable = False
                    for next_row in range(row + 2, len(self.datas)):
                        if next_row != len(self.datas) - 1:
                            if flag != self.datas[next_row][dictcol] and \
                                    self.datas[row][numcol] + self.datas[next_row][numcol] == 0:
                                remove_row.append(row)
                                remove_row.append(next_row)
                                removeEnable = True
                                break
                            if flag == self.datas[next_row][dictcol]:
                                break
                    if removeEnable:
                        break
            else:
                continue
        del_row = len(remove_row) - 1
        # print("待删除 : %s" % (remove_row))
        while del_row >= 0:
            del self.datas[remove_row[del_row]]
            del_row -= 1

        for row in range(len(self.datas)):
            if row != len(self.datas) - 1:
                flag = self.datas[row][dictcol]
                if flag != self.datas[row + 1][dictcol] and \
                        self.datas[row][numcol] + self.datas[row + 1][numcol] == 0:
                    return False
            else:
                continue
        return True

    def merge_trading_row(self):
        dictcol = self.get_col_by_tital('方向')
        amountcol = self.get_col_by_tital('成交金额')
        numcol = self.get_col_by_tital('成交数量')
        id = self.datas[0][self.get_col_by_tital('策略Id;策略ID;策略id')]

        self.direct = self.datas[0][dictcol]
        total = 0
        totalamount = 0
        for row in self.datas:
            total += row[numcol]
            totalamount += row[amountcol]

        count = 0
        while self._merge_trading() is False:
            # print(self.name + '检测合并' + str(count))
            count += 1

        totalamount2 = 0
        for row in self.datas:
            totalamount2 += row[amountcol]

        self.zerorow.append(id)
        self.zerorow.append(self.direct)
        self.zerorow.append(total)
        self.zerorow.append(totalamount)
        self.zerorow.append(totalamount2)
        self.zerorow.append(totalamount - totalamount2)


class TradingTable(TableData):
    def __init__(self, name, header=None, datas=None):
        super(TradingTable, self).__init__(name, header, datas)

    def merge_trading_row(self):
        remove_row = []
        dictcol = self.get_col_by_tital('方向')
        amountcol = self.get_col_by_tital('成交金额')
        numcol = self.get_col_by_tital('成交数量')
        self.direct = self.datas[0][dictcol]
        for row in range(len(self.datas)):
            if row in remove_row:
                continue
            if row != len(self.datas) - 1:
                flag = self.datas[row][dictcol]
                if flag != self.datas[row + 1][dictcol] and \
                        self.datas[row][numcol] + self.datas[row + 1][numcol] == 0:
                    remove_row.append(row)
                    remove_row.append(row + 1)
                    break
                elif flag != self.datas[row + 1][dictcol] and \
                        self.datas[row][numcol] + self.datas[row + 1][numcol] != 0:
                    for next_row in range(row + 2, len(self.datas)):
                        if next_row != len(self.datas) - 1:
                            if flag != self.datas[next_row][dictcol] and \
                                    self.datas[row][numcol] + self.datas[next_row][numcol] == 0:
                                remove_row.append(row)
                                remove_row.append(next_row)
                                break
                            if flag == self.datas[next_row][dictcol]:
                                break
            else:
                continue
        for del_row in remove_row:
            del self.datas[del_row]
        total = 0
        totalamount = 0
        for row in self.datas:
            total += row[numcol]
            totalamount += row[amountcol]
        self.zerorow.append(self.direct)
        self.zerorow.append(total)
        self.zerorow.append(totalamount)

    @staticmethod
    def gen_tradingtable(name, sheet):
        table = TradingTable(name)
        table.import_data_by_sheet(sheet)
        return table


def gen_summary(name, tables):
    datas = []
    headers = ["序号", "状态", "策略ID",
               "方向", "库存股数", "平均单价",
               "入库金额", "实现盈亏", "基准价",
               "档位价差", "买卖价差", "档位", "每档股数"
               ]
    sum_table = TableData(name + "-全")
    total = 0
    total_amount = 0
    avg_price = 0
    yl = 0

    row = 1
    for table in tables:
        total = total + table.zerorow[2]
        total_amount = total_amount + table.zerorow[4]
        yl = yl + table.zerorow[5]
        row_data = [row] + table.zerorow
        row_data.insert(1, "")
        for i in range(len(headers) - len(row_data)):
            row_data.append("")
        datas.append(row_data)
        row += 1
    # 股数不可为0
    if total != 0:
        avg_price = round(abs(total_amount / total), 3)
    else:
        avg_price = 0
    sum_table.set_datas(datas)
    sum_table.set_header(headers)
    sum_table.zerorow = ["合计", total, avg_price, total_amount, yl]
    
    total_header = ['序号','日期','时间','证券代码',
                    '证券简称','方向','成交数量','成交价格',
                    '成交金额','成交编号','委托编号',
                    '股东号','策略ID','组合ID','组合名称']
    #增加买入汇总
    buy_table = TableData("买入汇总")
    buy_table.zerorow.append(0)
    buy_table.direct = '买入'
    buy_datas = []
    for table in tables:
        if table.direct == '卖出':
            continue
        if table.zerorow[2] != 0:
            buy_datas += copy.deepcopy(table.datas)
            buy_table.zerorow[0] += table.zerorow[5]
    for i in range(len(buy_datas)):
        buy_datas[i][0] = i + 1
    buy_table.set_datas(buy_datas)
    buy_table.set_header(total_header)
    #增加卖出汇总
    sell_table = TableData("卖出汇总")
    sell_table.zerorow.append(0)
    sell_table.direct = '卖出'
    sell_datas = []
    for table in tables:
        if table.direct == '买入':
            continue
        if table.zerorow[2] != 0:
            sell_datas += copy.deepcopy(table.datas)
            sell_table.zerorow[0] += table.zerorow[5]
    for i in range(len(sell_datas)):
        sell_datas[i][0] = i + 1
    sell_table.set_datas(sell_datas)
    sell_table.set_header(total_header)
    res = [sum_table, buy_table, sell_table] + tables
    return res


def gen_sumtable(sheet, table):
    table.name = table.name.replace(' ', '')
    table.name = table.name.replace('-', '')
    # header
    row = 0
    col_width_list = []
    # title
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 18)))
    sheet.row(row).set_style(tall_style)
    #sheet.write_merge(row, row, 0, 1, "")
    sheet.write_merge(row, row, 0, len(table.header) -
                      1, table.name + "策略汇总表", get_excel_style('sumtable_tital'))

    row += 1
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 14)))
    sheet.row(row).set_style(tall_style)
    #sheet.write_merge(row, row, 0, 1, "")
    timestr = datetime.datetime.now().strftime(
        "%Y年1~") + str(int(datetime.datetime.now().strftime("%m"))) + "月"
    sheet.write_merge(row, row, 0, len(table.header) - 1, timestr, get_excel_style('sumtable_date'))

    row += 1
    tall_style = xlwt.easyxf('font: height ' + str(int(2 * 20 * 12)))
    sheet.row(row).set_style(tall_style)
    for col in range(len(table.zerorow)):
        style = get_excel_style('sumtable_zero')
        if col == 1:
            style.num_format_str = '#,##0'
        elif col == 2:
            style.num_format_str = '#,##0.000'
        elif col == 3 or col == 4:
            style.num_format_str = '#,##0.00'
        else:
            style.num_format_str = 'General'
            
        if col == 0:
            sheet.write(row, 0, '', style)
            sheet.write(row, 1, '', style)
            sheet.write(row, 2, '', style)
            col_width_list.append(0)
            col_width_list.append(0)
            col_width_list.append(0)
        sheet.write(row, col + 3, table.zerorow[col], style)
        # 计算字宽，如果没有文字，则不计算
        modify = 4
        if isinstance(table.zerorow[col], float):
            modify = 8
        col_wifth = (len(str(table.zerorow[col])) + modify) * 20 * 12
        col_width_list.append(col_wifth)
        # 填充不足单元格
        if col == len(table.zerorow) - 1:
            for new_col in range(len(table.header) - len(table.zerorow) - 3):
                sheet.write(row, col + 4 + new_col, '', style)
                col_width_list.append(0)

    row += 1
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 12)))
    sheet.row(row).set_style(tall_style)
    for col in range(len(table.header)):
        style = get_excel_style('sumtable_header')
        sheet.write(row, col, table.header[col], style)
        # 计算字宽，如果没有文字，则不计算
        col_wifth = (len(str(table.header[col])) + 10) * 20 * 12
        if col_wifth > col_width_list[col]:
            col_width_list[col] = col_wifth

    row += 1
    for r in table.datas:
        for col in range(len(r)):
            style = get_excel_style('sumtable_cell')
            if r[col] == "买入":
                style = get_excel_style('sumtable_cell_sell')
            elif r[col] == "卖出":
                style = get_excel_style('sumtable_cell_buy')
            else:
                style.font.colour_index = 0x08
            if col == 3 or col == 0:
                sumtable_cell_alignment = xlwt.Alignment()
                sumtable_cell_alignment.horz = 0x02
                sumtable_cell_alignment.vert = 0x01
                style.alignment = sumtable_cell_alignment
            else:
                sumtable_cell_alignment = xlwt.Alignment()
                sumtable_cell_alignment.horz = 0x03
                sumtable_cell_alignment.vert = 0x01
                style.alignment = sumtable_cell_alignment
                #style.alignment.horz = 0x03
            if isinstance(r[col], float):
                if col == 5:
                    if r[4] != 0:
                        r[col] = abs(r[6] / r[4])
                    else:
                        r[col] = 0
                    style.num_format_str = '#,##0.000'
                elif col == 4:
                    style.num_format_str = '#,##0'
                else:
                    style.num_format_str = '#,##0.00'
                sheet.write(row, col, r[col], style)
            else:
                style.num_format_str = 'General'
                sheet.write(row, col, r[col], style)
            # 计算字宽，如果没有文字，则不计算
            col_wifth = (len(str(r[col])) + 6) * 20 * 12
            if col_wifth > col_width_list[col]:
                col_width_list[col] = col_wifth
        tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 12)))
        sheet.row(row).set_style(tall_style)
        row += 1

    #填充公式
    for col in range(len(table.header)):
        style = get_excel_style('sumtable_total')
        if col == 4:
            style.num_format_str = '#,##0'
        elif col == 5:
            style.num_format_str = '#,##0.000'
        elif col == 7 or col == 6:
            style.num_format_str = '#,##0.00'
        else:
            style.num_format_str = 'General'
        if col == 5:
            sheet.write(row,col,xlwt.Formula('ABS(G' + str(row + 1) + ')/ABS(E' + str(row + 1) + ')'), style)
        elif col == 6:
            sheet.write(row,col,xlwt.Formula('SUM(G5:'+'G' + str(row) + ')'), style)
        elif col == 7:
            sheet.write(row,col,xlwt.Formula('SUM(H5:'+'H' + str(row) + ')'), style)
        elif col == 4:
            sheet.write(row,col,xlwt.Formula('SUM(E5:'+'E' + str(row) + ')'), style)
        elif col == 3:
            sheet.write(row,col,'求和', style)
        else:
            sheet.write(row,col,'', style)
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 12)))
    sheet.row(row).set_style(tall_style)

    # 设置列宽
    for col in range(len(col_width_list)):
        colfnt = sheet.col(col)
        colfnt.width = col_width_list[col]

def gen_totaltable(sheet, table):
    row = 0
    col_width_list = []
    for col in range(len(table.header)):
        style_name = 'totaltable_zero_'
        flag = 'sell' if table.direct != '买入' else 'buy'
        style = get_excel_style(style_name + flag)

        if col == 6:
            style.num_format_str = '#,##0'
        elif col == 7:
            style.num_format_str = '#,##0.000'
        elif col == 8 or col == 9:
            style.num_format_str = '#,##0.00'
        else:
            style.num_format_str = 'General'
        if col <= 3 or col > 8:
            sheet.write(row, col, '', style)
            col_width_list.append(0)
        elif col == 4:
            sheet.write(row, col, "合计", style)
        elif col == 5:
            sheet.write(row, col, table.direct, style)
        elif col == 6:
            sheet.write(row,col,xlwt.Formula('SUM(G3:'+'G' + str(len(table.datas) + 2) + ')'), style)
        elif col == 7:
            sheet.write(row,col,xlwt.Formula('ABS(I1)/ABS(G1)'), style)
        elif col == 8:
            sheet.write(row,col,xlwt.Formula('SUM(I3:'+'I' + str(len(table.datas) + 2) + ')'), style)
        #elif col == 9:
        #    sheet.write(row,col,table.zerorow[0], style)
    row += 1
    for col in range(len(table.header)):
        style = get_excel_style('totaltable_header')
        sheet.write(row, col, table.header[col], style)
        # 计算字宽，如果没有文字，则不计算
        col_wifth = (len(str(table.header[col])) + 12) * 20 * 14
        col_width_list.append(col_wifth)
    row += 1

    for r in table.datas:
        for col in range(len(r)):
            style = get_excel_style('totaltable_data')
            if isinstance(r[col], float):
                totaltable_cell_alignment = xlwt.Alignment()
                totaltable_cell_alignment.horz = 0x03
                totaltable_cell_alignment.vert = 0x01
                style.alignment = totaltable_cell_alignment
                #style.alignment.horz = 0x03
                if col == deal_price_col:
                    style.num_format_str = '#,##0.000'
                elif col == deal_num_col:
                    style.num_format_str = '#,##0'
                else:
                    style.num_format_str = '#,##0.00'
                sheet.write(row, col, r[col], style)
            else:
                totaltable_cell_alignment = xlwt.Alignment()
                totaltable_cell_alignment.horz = 0x02
                totaltable_cell_alignment.vert = 0x01
                style.alignment = totaltable_cell_alignment
                style.num_format_str = 'General'
                sheet.write(row, col, r[col], style)
            # 计算字宽，如果没有文字，则不计算
            col_wifth = (len(str(r[col])) + 12) * 20 * 12
            if col_wifth > col_width_list[col]:
                col_width_list[col] = col_wifth
        row += 1
    # 设置列宽
    for col in range(len(col_width_list)):
        colfnt = sheet.col(col)
        colfnt.width = col_width_list[col]

def gen_strategytable(sheet, table):
    # header
    row = 0
    col_width_list = []

    for col in range(len(table.zerorow)):
        style_name = 'strategy_zero_'
        directFlagCount = 1
        flag = 'sell' if table.zerorow[directFlagCount] != '买入' else 'buy'
        style = get_excel_style(style_name + flag)

        if col == 2:
            style.num_format_str = '#,##0'
        elif col == 3:
            if table.zerorow[2] != 0:
                table.zerorow[col] = abs(
                    table.zerorow[4] / table.zerorow[2])
            else:
                table.zerorow[col] = 0
            style.num_format_str = '#,##0.000'
        elif col == 4 or col == 5:
            style.num_format_str = '#,##0.00'
        else:
            style.num_format_str = 'General'
        if col == 0:
            sheet.write(row, 0, '', style)
            sheet.write(row, 1, '', style)
            sheet.write(row, 2, '', style)
            sheet.write(row, 3, '', style)
            col_width_list.append(0)
            col_width_list.append(0)
            col_width_list.append(0)
            col_width_list.append(0)
        sheet.write(row, col + 4, table.zerorow[col], style)

        # 计算字宽，如果没有文字，则不计算
        modify = 4
        if isinstance(table.zerorow[col], float):
            modify = 6
        col_wifth = (len(str(table.zerorow[col])) + modify) * 20 * 14
        col_width_list.append(col_wifth)
        # 填充不足单元格
        if col == len(table.zerorow) - 1:
            for new_col in range(len(table.header) - len(table.zerorow) - 3):
                sheet.write(row, col + 5 + new_col, '', style)
                col_width_list.append(0)
    row += 1
    for col in range(len(table.header)):
        style = get_excel_style('strategy_header')
        sheet.write(row, col, table.header[col], style)
        # 计算字宽，如果没有文字，则不计算
        col_wifth = (len(str(table.header[col])) + 4) * 20 * 14
        if col_wifth > col_width_list[col]:
            col_width_list[col] = col_wifth
    row += 1

    for r in table.datas:
        for col in range(len(r)):
            style = get_excel_style('strategy_data')
            if isinstance(r[col], float):
                strategytable_cell_alignment = xlwt.Alignment()
                strategytable_cell_alignment.horz = 0x03
                strategytable_cell_alignment.vert = 0x01
                style.alignment = strategytable_cell_alignment
                #style.alignment.horz = 0x03
                if col == deal_price_col:
                    style.num_format_str = '#,##0.000'
                elif col == deal_num_col:
                    style.num_format_str = '#,##0'
                else:
                    style.num_format_str = '#,##0.00'
                sheet.write(row, col, r[col], style)
            else:
                strategytable_cell_alignment = xlwt.Alignment()
                strategytable_cell_alignment.horz = 0x02
                strategytable_cell_alignment.vert = 0x01
                style.alignment = strategytable_cell_alignment
                style.num_format_str = 'General'
                sheet.write(row, col, r[col], style)
            # 计算字宽，如果没有文字，则不计算
            col_wifth = (len(str(r[col])) + 4) * 20 * 11
            if col_wifth > col_width_list[col]:
                col_width_list[col] = col_wifth
        row += 1
    # 设置列宽
    for col in range(len(col_width_list)):
        colfnt = sheet.col(col)
        colfnt.width = col_width_list[col]


def parse_excel(excel, filter=['成交', '委托', '持仓', '资金']):
    res = {}
    tables = {}
    for sn in excel.sheet_names():
        if sn in filter:
            continue
        else:
            tables[sn] = excel.sheet_by_name(sn)
    print(tables)
    for k, v in tables.items():
        table = TradingTable.gen_tradingtable(k, v)
        tabledatas = table.split_table_by_tital('策略Id;策略ID;策略id')
        print(tabledatas)
        for t in tabledatas:
            pricecol = t.get_col_by_tital('成交价格')
            amountcol = t.get_col_by_tital('成交金额')
            numcol = t.get_col_by_tital('成交数量')

            t.fomat_col([pricecol, amountcol, numcol], float)
            t.merge_row_by_tital(['委托编号', '日期'], [amountcol, numcol])
            t.show_datas()
            t.merge_trading_row()
            t.reset_number()
        res[k] = gen_summary(k, tabledatas)
    return res


def gen_total_excel(tabledatas, outfile):
    for k, tables in tabledatas.items():
        workbook = xlwt.Workbook(encoding='utf8')
        xlwt.add_palette_colour('sell color', 0x21)
        xlwt.add_palette_colour('buy color', 0x22)
        xlwt.add_palette_colour('sell text color', 0x23)
        xlwt.add_palette_colour('buy text color', 0x24)
        xlwt.add_palette_colour('total background color', 0x25)
        xlwt.add_palette_colour('total buy text color', 0x26)
        xlwt.add_palette_colour('total sell text color', 0x27)
        xlwt.add_palette_colour('sum data color', 0x28)
        workbook.set_colour_RGB(0x21, 255, 0, 0)
        workbook.set_colour_RGB(0x22, 0, 176, 80)
        workbook.set_colour_RGB(0x23, 0, 32, 96)
        workbook.set_colour_RGB(0x24, 255, 255, 0)
        workbook.set_colour_RGB(0x25, 155, 194, 230)
        workbook.set_colour_RGB(0x26, 255, 0, 0)
        workbook.set_colour_RGB(0x27, 0, 176, 80)
        workbook.set_colour_RGB(0x28, 255, 255, 0)

        for table in tables:
            sheet = workbook.add_sheet(table.name, cell_overwrite_ok=True)
            if "-全" in table.name:
                gen_sumtable(sheet, table)
            elif '买入汇总' in table.name or '卖出汇总' in table.name:
                gen_totaltable(sheet, table)
            else:
                gen_strategytable(sheet, table)

        date = datetime.datetime.now().strftime("%Y-%m-%d")
        workbook.save(outfile + '/' + k + "_" + date + "汇总表.xlsx")


def merge_excel(file, outfile):
    # outfile 改为 保存路径
    impexcel = xlrd.open_workbook(file)
    tables = parse_excel(impexcel)
    gen_total_excel(tables, outfile)
