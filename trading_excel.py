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

import datetime
try:
    import xlrd
    import xlwt
except Exception as e:
    print(e)

deal_num_col = 5
deal_price_col = 6
deal_amount_col = 7


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
    sum_table = TableData(name + "-汇总")
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
    res = [sum_table] + tables
    return res


def gen_sumtable(sheet, table):
    table.name = table.name.replace(' ', '')
    table.name = table.name.replace('-', '')
    # header
    row = 0
    col_width_list = []
    # title
    style = xlwt.XFStyle()  # Create the pattern
    tfont = xlwt.Font()
    tfont.name = u"微软黑雅"
    tfont.bold = True
    tfont.height = 20*18
    style.font = tfont
    talignment = xlwt.Alignment()
    talignment.horz = 0x02
    talignment.vert = 0x01
    style.alignment = talignment
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 18)))
    sheet.row(row).set_style(tall_style)
    sheet.write_merge(row, row, 0, 1, "")
    sheet.write_merge(row, row, 2, len(table.header) -
                      1, table.name + "策略汇总表", style)

    row += 1
    style = xlwt.XFStyle()  # Create the pattern
    tfont = xlwt.Font()
    tfont.name = u"等线"
    tfont.bold = True
    tfont.height = 20*14
    style.font = tfont
    talignment = xlwt.Alignment()
    talignment.horz = 0x02
    talignment.vert = 0x01
    style.alignment = talignment
    tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 14)))
    sheet.row(row).set_style(tall_style)
    sheet.write_merge(row, row, 0, 1, "")
    timestr = datetime.datetime.now().strftime(
        "%Y年1~") + str(int(datetime.datetime.now().strftime("%m"))) + "月"
    sheet.write_merge(row, row, 2, len(table.header) - 1, timestr, style)

    row += 1
    tall_style = xlwt.easyxf('font: height ' + str(int(2 * 20 * 12)))
    sheet.row(row).set_style(tall_style)
    for col in range(len(table.zerorow)):
        style = xlwt.XFStyle()  # Create the pattern
        pattern = xlwt.Pattern()  # Create the pattern
        # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 0x25
        font = xlwt.Font()
        font.name = u"宋体"
        font.bold = True
        font.height = 20*12
        style.font = font

        alignment = xlwt.Alignment()
        alignment.horz = 0x03
        alignment.vert = 0x01
        style.alignment = alignment
        if col == 1:
            style.num_format_str = '#,##0'
        elif col == 2:
            style.num_format_str = '#,##0.000'
        elif col == 3 or col == 4:
            style.num_format_str = '#,##0.00'
        style.pattern = pattern  # Add pattern to style
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
        style = xlwt.XFStyle()  # Create the pattern
        font = xlwt.Font()
        font.name = u"宋体"
        font.bold = True
        font.height = 20*12
        style.font = font

        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        borders.left_colour = 0xff
        borders.right_colour = 0xff
        borders.top_colour = 0xff
        borders.bottom_colour = 0xff
        style.borders = borders

        alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment.vert = 0x01
        style.alignment = alignment
        sheet.write(row, col, table.header[col], style)
        # 计算字宽，如果没有文字，则不计算
        col_wifth = (len(str(table.header[col])) + 10) * 20 * 12
        if col_wifth > col_width_list[col]:
            col_width_list[col] = col_wifth

    row += 1
    for r in table.datas:
        for col in range(len(r)):
            style = xlwt.XFStyle()

            borders = xlwt.Borders()
            borders.left = xlwt.Borders.THIN
            borders.right = xlwt.Borders.THIN
            borders.top = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            borders.left_colour = 0xff
            borders.right_colour = 0xff
            borders.top_colour = 0xff
            borders.bottom_colour = 0xff
            style.borders = borders

            font = xlwt.Font()
            font.name = u"宋体"
            font.height = 20*12
            if r[col] == "买入":
                font.colour_index = 0x26
            elif r[col] == "卖出":
                font.colour_index = 0x27
            style.font = font

            alignment = xlwt.Alignment()
            alignment.horz = 0x03
            alignment.vert = 0x01
            style.alignment = alignment
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
                sheet.write(row, col, r[col], style)
            # 计算字宽，如果没有文字，则不计算
            col_wifth = (len(str(r[col])) + 6) * 20 * 12
            if col_wifth > col_width_list[col]:
                col_width_list[col] = col_wifth
        tall_style = xlwt.easyxf('font: height ' + str(int(1.5 * 20 * 12)))
        sheet.row(row).set_style(tall_style)
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
        style = xlwt.XFStyle()  # Create the pattern
        pattern = xlwt.Pattern()  # Create the pattern
        # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        directFlagCount = 1
        # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon,
        # 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        pattern.pattern_fore_colour = 0x22 if table.zerorow[directFlagCount] != '买入' else 0x21

        font = xlwt.Font()
        font.name = u"等线"
        font.colour_index = 0x23 if table.zerorow[directFlagCount] != '买入' else 0x24
        font.bold = True
        font.height = 20*14
        style.font = font

        alignment = xlwt.Alignment()
        alignment.horz = 0x03
        alignment.vert = 0x01
        style.alignment = alignment
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
        style.pattern = pattern  # Add pattern to style
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
        style = xlwt.XFStyle()  # Create the pattern
        font = xlwt.Font()
        font.name = u"等线"
        font.bold = True
        font.height = 20*12
        style.font = font

        alignment = xlwt.Alignment()
        # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.horz = 0x02
        # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment.vert = 0x01
        style.alignment = alignment
        sheet.write(row, col, table.header[col], style)
        # 计算字宽，如果没有文字，则不计算
        col_wifth = (len(str(table.header[col])) + 4) * 20 * 14
        if col_wifth > col_width_list[col]:
            col_width_list[col] = col_wifth
    row += 1

    for r in table.datas:
        for col in range(len(r)):
            style = xlwt.XFStyle()
            font = xlwt.Font()
            font.name = u"等线"
            font.height = 20*11
            style.font = font

            if isinstance(r[col], float):
                if col == deal_price_col:
                    style.num_format_str = '#,##0.000'
                elif col == deal_num_col:
                    style.num_format_str = '#,##0'
                else:
                    style.num_format_str = '#,##0.00'
                sheet.write(row, col, r[col], style)
            else:
                alignment = xlwt.Alignment()
                alignment.horz = 0x03
                alignment.vert = 0x01
                style.alignment = alignment
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
        workbook.set_colour_RGB(0x21, 255, 0, 0)
        workbook.set_colour_RGB(0x22, 0, 176, 80)
        workbook.set_colour_RGB(0x23, 0, 32, 96)
        workbook.set_colour_RGB(0x24, 255, 255, 0)
        workbook.set_colour_RGB(0x25, 155, 194, 230)
        workbook.set_colour_RGB(0x26, 255, 0, 0)
        workbook.set_colour_RGB(0x27, 0, 176, 80)

        for table in tables:
            sheet = workbook.add_sheet(table.name, cell_overwrite_ok=True)
            if "-汇总" in table.name:
                gen_sumtable(sheet, table)
            else:
                gen_strategytable(sheet, table)

        date = datetime.datetime.now().strftime("%Y-%m-%d")
        workbook.save(outfile + '/' + k + "_" + date + "汇总表.xlsx")


def merge_excel(file, outfile):
    # outfile 改为 保存路径
    impexcel = xlrd.open_workbook(file)
    tables = parse_excel(impexcel)
    gen_total_excel(tables, outfile)
