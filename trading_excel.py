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

try :
    import xlrd, xlwt
except Exception as e:
    print(e)

deal_num_col = 5
deal_price_col = 6
deal_amount_col = 7

class TableData:
    def __init__(self, name, header = None,datas = None):
        self.name = name
        self.zerorow = []
        self.header = header
        self.datas = datas

    def import_data_by_sheet(self, sheet):
        self.datas = []
        for i in range(sheet.nrows):
            if i == 0:
                self.header = sheet.row_values(i)
            else:
                self.datas.append(sheet.row_values(i))

    def set_header(self, header):
        self.header = header

    def set_datas(self, datas):
        self.datas = datas

    def fomat_col(self,col_lsit, t):
        for i in range(len(self.datas)):
            for c in col_lsit:
                v = self.datas[i][c]
                if t == float or t == int:
                    s = ''
                    if isinstance(v,str) and ',' in v:
                        vs = v.split(',')
                        for sub in vs:
                            s += sub
                        v = s
                self.datas[i][c] = t(v)

    def get_col_by_tital(self, tital):
        for i in range(len(self.header)):
            if self.header[i] == tital:
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
        for k,v in keymap.items():
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
        if self.datas is None: return None
        for row in self.datas:
            if row[col] not in keymap.keys():
                keymap[row[col]] = []
            keymap[row[col]].append(row)
        res = []
        for k,v in keymap.items():
            table = TableData(self.name + '-' + k[len(k)-9:])
            table.set_header(self.header)
            table.set_datas(v)
            res.append(table)
        return res

    def split_table_by_tital(self,tital):
        for i in range(len(self.header)):
            if self.header[i] == tital:
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
                #print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                #print(flag)
                #print(self.datas[row + 1][dictcol])
                #print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
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
        #print("待删除 : %s" % (remove_row))
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
        id = self.datas[0][self.get_col_by_tital('策略Id')]

        self.direct = self.datas[0][dictcol]
        total = 0
        totalamount = 0
        for row in self.datas:
            total += row[numcol]
            totalamount += row[amountcol]

        count = 0
        while self._merge_trading() is False:
            #print(self.name + '检测合并' + str(count))
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
        buy = '买入'
        sell = '卖出'
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

def parse_excel(excel, filter = ['成交', '委托', '持仓', '资金']):
    res = []
    tables = {}
    for sn in excel.sheet_names():
        if sn in filter:
            continue
        else:
            tables[sn] = excel.sheet_by_name(sn)
    for k,v in tables.items():
        table = TradingTable.gen_tradingtable(k,v)
        tabledatas = table.split_table_by_tital('策略Id')
        for t in tabledatas:
            t.fomat_col([5,6,7],float)
            t.merge_row_by_tital(['委托编号','日期'],[5,7])
            t.show_datas()
        res += tabledatas
    return res

def gen_total_excel(tabledatas, outfile):
    workbook = xlwt.Workbook(encoding='utf8')
    for table in tabledatas:
        table.merge_trading_row()
        sheet = workbook.add_sheet(table.name)
        #header
        row = 0
        for col in range(len(table.zerorow)):
            pattern = xlwt.Pattern()  # Create the pattern
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
            directFlagCount = 1
            # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 
            # 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
            pattern.pattern_fore_colour = 2 if table.zerorow[directFlagCount] != '卖出' else 3

            style = xlwt.XFStyle()  # Create the pattern
            alignment = xlwt.Alignment()
            alignment.horz = 0x03
            style.alignment = alignment
            if col == 2:
                style.num_format_str = '#,##0'
            elif col == 3:
                if table.zerorow[2] != 0:
                    table.zerorow[col] = abs(table.zerorow[4] / table.zerorow[2])
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
            sheet.write(row, col + 3, table.zerorow[col], style)

        row += 1
        for col in range(len(table.header)):
            sheet.write(row, col, table.header[col])
        row += 1

        for r in table.datas:
            for col in range(len(r)):
                if isinstance(r[col], float):
                    style = xlwt.XFStyle()
                    if col == deal_price_col:
                        style.num_format_str = '#,##0.000'
                    elif col == deal_num_col:
                        style.num_format_str = '#,##0'
                    else:
                        style.num_format_str = '#,##0.00'
                    sheet.write(row, col, r[col], style)
                else:
                    sheet.write(row, col, r[col])
            row += 1
    workbook.save(outfile)

def merge_excel(file, outfile):
    impexcel = xlrd.open_workbook(file)
    tables = parse_excel(impexcel)
    gen_total_excel(tables, outfile)
