# coding=utf-8
import xlwings as xw
import os
import sys


def wblink():
    # 获取当前路径
    _path_c = os.getcwd()
    # 初始化文件名
    _xlsx_name = ''

    for _file_name in os.listdir(_path_c):
        if _file_name[-4:] == "xlsx" and _file_name[0] != '~':
            _xlsx_name = _file_name
            break
    if _xlsx_name == '':
        print('没有找到Excel表格,请检查当前目录:' + _path_c)
        sys.exit(0)

    # 链接工作簿
    workbook = xw.Book(os.path.join(_path_c, _xlsx_name))
    workbook.api.Application.ErrorCheckingOptions.BackgroundChecking = False  # 关闭Excel错误检查
    # 检查有资产负债表存在
    _sheet_index = 1
    _bs_name = ''
    while _sheet_index <= len(workbook.sheets):
        if workbook.api.Worksheets(_sheet_index).Name == '资产负债表':
            _bs_name = '资产负债表'
            break
        _sheet_index = _sheet_index + 1
    if _bs_name == '':
        print('没有找到资产负债表,请检查Excel文件是否正确:' + _xlsx_name)
        sys.exit(0)
    return workbook


wb = wblink()


# 创建Sheets
def new_sheet(x, y):
    _sheet_index = 1
    _sheet_exist = 0
    for _sheet_name in wb.sheets:
        if wb.api.Worksheets(_sheet_index).Name == x:
            _sheet_exist = 1
        _sheet_index = _sheet_index + 1
    if _sheet_exist == 0:
        wb.sheets.add(x, after=y)
    else:
        wb.sheets(x).clear()
    return


# 获取公司名称和日期
def get_date_name():
    _date = ""
    _name = ""
    for _date in wb.sheets('资产负债表').range('A1:A5').value:
        if str(_date)[-1] == '日':
            break

    for _name in wb.sheets('资产负债表').range('B1:B5').value:
        if str(_name)[-2:] == '公司' or str(_name)[-2:] == '.,':
            break
    date_name = [str(_date), str(_name)]
    return date_name


# -------------------------------preBS Start-----------------------------------------
def fill_prebs():
    # preBS
    st = [['资产负债表', '', '', '', '', ''], ['编制单位:', '', '', '', '', '报表日期:'],
          [get_date_name()[1], '', '', '', '', get_date_name()[0]], ['资产:', '期末数', '年初数', '负债和股东权益:', '期末数', '年初数']]
    for _column in [['流动资产：'], ['货币资金'], ['交易性金融资产'], ['衍生金融资产'], ['应收票据'], ['应收账款'], ['应收款项融资'], ['预付款项'], ['其它应收款'],
                    ['其中：应收利息'], ['应收股利'], ['存货'], ['合同资产'], ['划分为持有待售的资产'], ['一年内到期的非流动资产'], ['其它流动资产'], ['流动资产合计'],
                    ['非流动资产：'], ['发放贷款和垫款'], ['债权投资'], ['其它债权投资'], ['长期应收款'], ['长期股权投资'], ['其它权益工具投资'], ['其它非流动金融资产'],
                    ['投资性房地产'], ['固定资产'], ['在建工程'], ['生产性生物资产'], ['油气资产'], ['使用权资产'], ['无形资产'], ['开发支出'], ['商誉'],
                    ['长期待摊费用'], ['递延所得税资产'], ['其它非流动资产'], [''], [''], [''], [''], [''], [''], [''], [''], [''],
                    ['非流动资产合计'], ['资产总计']]:
        st.append(_column)
        for _x in range(5):
            st[st.index(_column)].append('')

    _index = st.index(['流动资产：', '', '', '', '', ''])
    for _column in ['流动负债：', '短期借款', '交易性金融负债', '衍生金融负债', '应付票据', '应付账款', '预收款项', '合同负债', '应付职工薪酬', '应交税费', '其它应付款',
                    '其中：应付利息', '应付股利', '划分为持有待售的负债', '一年内到期的非流动负债', '其它流动负债', '流动负债合计', '非流动负债：', '保险合同准备金', '长期借款',
                    '应付债券', '其中：优先股(负债)', '永续债(负债)', '租赁负债', '长期应付款', '长期应付职工薪酬', '预计负债', '递延收益', '递延所得税负债', '其它非流动负债',
                    '非流动负债合计', '负债合计', '所有者权益：', '实收资本(股本)', '其它权益工具', '其中：优先股(权益)', '永续债(权益)', '资本公积', '减：库存股',
                    '其它综合收益', '专项储备', '盈余公积', '一般风险准备', '未分配利润', '归属于母公司所有者权益合计', '少数股东权益', '股东权益合计', '负债和股东权益总计']:
        st[_index][3] = _column
        _index = _index + 1
        _num_table = gen_list('资产负债表', 6)
    match_1(_num_table, st)
    match_1(_num_table, st, 3)
    match_2(_num_table, '应收利息', st, '其中：应收利息')
    match_2(_num_table, '可供出售金融资产', st, '其它非流动金融资产')
    match_2(_num_table, '工程物资', st, '在建工程')
    match_2(_num_table, '固定资产清理', st, '固定资产')
    match_2(_num_table, '应付利息', st, '其中：应付利息', 3, 3)
    match_2(_num_table, '应收利息', st, '其它应收款')
    match_2(_num_table, '应收股利', st, '其它应收款')
    match_2(_num_table, '应付利息', st, '其它应付款', 3, 3)
    match_2(_num_table, '应付股利', st, '其它应付款', 3, 3)
    match_3(st, '流动资产合计', '=sum(B6:B20)-B14-B15')  # 流动资产合计数,去掉'其中'部分
    match_3(st, '流动资产合计', '=sum(C6:C20)-C14-C15', 0, 2)
    match_3(st, '非流动资产合计', '=sum(B23:B41)')  # 非流动资产
    match_3(st, '非流动资产合计', '=sum(C23:C41)', 0, 2)
    match_3(st, '资产总计', '=sum(B21,B51)')  # 资产
    match_3(st, '资产总计', '=sum(C21,C51)', 0, 2)
    match_3(st, '流动负债合计', '=sum(E6:E20)-E16-E17', 3, 1)  # 流动负债合计数,去掉'其中'部分
    match_3(st, '流动负债合计', '=sum(F6:F20)-F16-F17', 3, 2)
    match_3(st, '非流动负债合计', '=sum(E23:E34)-E26-E27', 3, 1)  # 非流动负债,去掉'其中'部分
    match_3(st, '非流动负债合计', '=sum(F23:F34)-F26-F27', 3, 2)
    match_3(st, '负债合计', '=sum(E21,E24)', 3, 1)  # 负债
    match_3(st, '负债合计', '=sum(F21,F24)', 3, 2)
    match_3(st, '归属于母公司所有者权益合计', '=sum(E38,E39,E42,-E43,E44:E48)', 3, 1)  # 归母权益
    match_3(st, '归属于母公司所有者权益合计', '=sum(F38,F39,F42,-F43,F44:F48)', 3, 2)
    match_3(st, '股东权益合计', '=sum(E49,E50)', 3, 1)  # 所有者权益
    match_3(st, '股东权益合计', '=sum(F49,F50)', 3, 2)
    match_3(st, '负债和股东权益总计', '=sum(E36,E51)', 3, 1)  # 负债和所有者权益
    match_3(st, '负债和股东权益总计', '=sum(F36,F51)', 3, 2)
    _j = 0
    for _list in st:
        _i = -1
        for _value in _list:
            _i = _i + 1
            if isinstance(_value, str):
                _list[_i] = _value.replace('其它', '其他')
        st[_j] = _list
        _j = _j + 1
    return st


def gen_list(sheet_name, start_row):
    num_table = wb.sheets(sheet_name).range('A' + str(start_row) + ':F80').value
    for _i in range(81 - start_row):
        for _j in range(6):
            if num_table[_i][_j] is None:
                num_table[_i][_j] = ''
            elif isinstance(num_table[_i][_j], str):
                num_table[_i][_j] = num_table[_i][_j].replace(' ', '')
                num_table[_i][_j] = num_table[_i][_j].replace('\u3000', '')
    # print('nt:', num_table)
    return num_table


def match_1(num_table, st, col=0):
    for _match1 in num_table:
        for _match2 in st:
            if _match1[col] == _match2[col] and _match1[col] != '':
                # print(st[st.index(_match2)])
                st[st.index(_match2)][col + 1] = num_table[num_table.index(_match1)][col + 1]
                st[st.index(_match2)][col + 2] = num_table[num_table.index(_match1)][col + 2]
                break
    # print('st:', st)
    return


def match_2(num_table, value_1, st, value_2, col_1=0, col_2=0, nega=1):
    for _list1 in num_table:
        if _list1[col_1] == value_1:
            for _list2 in st:
                if _list2[col_2] == value_2:
                    if st[st.index(_list2)][col_2 + 1] == '':
                        st[st.index(_list2)][col_2 + 1] = num_table[num_table.index(_list1)][col_1 + 1] * nega
                    else:
                        st[st.index(_list2)][col_2 + 1] = float(st[st.index(_list2)][col_2 + 1]) + \
                                                          num_table[num_table.index(_list1)][col_1 + 1] * nega

                    if st[st.index(_list2)][col_2 + 2] == '':
                        st[st.index(_list2)][col_2 + 2] = num_table[num_table.index(_list1)][col_1 + 2] * nega
                        break
                    else:
                        st[st.index(_list2)][col_2 + 2] = float(st[st.index(_list2)][col_2 + 2]) + \
                                                          num_table[num_table.index(_list1)][col_1 + 2] * nega
                        break
    # print('st2:', st)
    return


def match_3(st, head, value, col_1=0, col_2=1):
    _i = -1
    for _list in st:
        _i = _i + 1
        # print(_list)
        if _list[col_1] == head:
            st[_i][col_1 + col_2] = value
            break
    return


# 数据格式设置
def format_prebs():
    wb.sheets('preBS').range('A1:F80').api.Font.Size = 11
    wb.sheets('preBS').range('A1:A1').api.ColumnWidth = 30
    wb.sheets('preBS').range('D1:D1').api.ColumnWidth = 30
    wb.sheets('preBS').range('B1:C1').api.ColumnWidth = 23
    wb.sheets('preBS').range('E1:F1').api.ColumnWidth = 23
    wb.sheets('preBS').range('A:F').api.Font.Name = "Arial"
    wb.sheets('preBS').range('A1:F80').api.Style = "Comma"
    wb.sheets('preBS').range('A4:F52').api.Borders(7).LineStyle = 1
    wb.sheets('preBS').range('A4:F52').api.Borders(8).LineStyle = 1
    wb.sheets('preBS').range('A4:F52').api.Borders(9).LineStyle = 1
    wb.sheets('preBS').range('A4:F52').api.Borders(10).LineStyle = 1
    wb.sheets('preBS').range('A4:F52').api.Borders(11).LineStyle = 1
    wb.sheets('preBS').range('A4:F52').api.Borders(12).LineStyle = 1
    wb.sheets('preBS').range('A1:F1').api.HorizontalAlignment = 7
    wb.sheets('preBS').range('A21:F21').api.Font.Bold = True
    wb.sheets('preBS').range('D35:F36').api.Font.Bold = True
    wb.sheets('preBS').range('A51:F52').api.Font.Bold = True
    for lst in range(15):
        wb.sheets('preBS').range('A' + str(lst + 6)).api.IndentLevel = 1
    for lst in [14, 15]:
        wb.sheets('preBS').range('A' + str(lst)).api.IndentLevel = 2
    for lst in range(19):
        wb.sheets('preBS').range('A' + str(lst + 23)).api.IndentLevel = 1
    for lst in range(15):
        wb.sheets('preBS').range('D' + str(lst + 6)).api.IndentLevel = 1
    for lst in [16, 17]:
        wb.sheets('preBS').range('D' + str(lst)).api.IndentLevel = 2
    for lst in range(12):
        wb.sheets('preBS').range('D' + str(lst + 23)).api.IndentLevel = 1
    for lst in [26, 27]:
        wb.sheets('preBS').range('D' + str(lst)).api.IndentLevel = 2
    for lst in range(11):
        wb.sheets('preBS').range('D' + str(lst + 38)).api.IndentLevel = 1
    for lst in [40, 41]:
        wb.sheets('preBS').range('D' + str(lst)).api.IndentLevel = 2

    return


# ----------------------------preBS Over-------------------------------------------
# ----------------------------preIS Start-------------------------------------------
def fill_preis():
    st = [['利润表', '', ''], ['编制单位:', '', '报表日期:'], [get_date_name()[1], '', get_date_name()[0]],
          ['项目:', '本期金额', '本年累计']]
    for _column in [['一、营业总收入'], ['其中：营业收入'], ['利息收入_'], ['已赚保费'], ['手续费及佣金收入'], ['二、营业总成本'], ['其中：营业成本'], ['利息支出'],
                    ['手续费及佣金支出'], ['退保金'], ['赔付支出净额'], ['提取保险责任准备金净额'], ['保单红利支出'], ['分保费用'], ['税金及附加'], ['销售费用'],
                    ['管理费用'], ['研发费用'], ['财务费用'], ['其中：利息费用'], ['利息收入'], ['加：其他收益'], ['投资收益（损失以“－”号填列）'],
                    ['其中：对联营企业和合营企业的投资收益'], ['以摊余成本计量的金融资产终止确认收益'], ['汇兑收益（损失以“-”号填列）'], ['净敞口套期收益（损失以“－”号填列）'],
                    ['公允价值变动收益（损失以“－”号填列）'], ['信用减值损失（损失以“-”号填列）'], ['资产减值损失（损失以“-”号填列）'], ['资产处置收益（损失以“-”号填列）'],
                    ['三、营业利润（亏损以“－”号填列）'], ['加：营业外收入'], ['减：营业外支出'], ['四、利润总额（亏损总额以“－”号填列）'], ['减：所得税费用'],
                    ['五、净利润（净亏损以“－”号填列）'], ['（一）按经营持续性分类'], ['1.持续经营净利润（净亏损以“－”号填列）'], ['2.终止经营净利润（净亏损以“－”号填列）'],
                    ['（二）按所有权归属分类'], ['1.归属于母公司所有者的净利润'], ['2.少数股东损益'], ['六、其他综合收益的税后净额'], ['归属母公司所有者的其他综合收益的税后净额'],
                    ['（一）不能重分类进损益的其他综合收益'], ['1.重新计量设定受益计划变动额'], ['2.权益法下不能转损益的其他综合收益'], ['3.其他权益工具投资公允价值变动'],
                    ['4.企业自身信用风险公允价值变动'], ['5.其他'], ['（二）将重分类进损益的其他综合收益'], ['1.权益法下可转损益的其他综合收益'], ['2.其他债权投资公允价值变动'],
                    ['3.金融资产重分类计入其他综合收益的金额'], ['4.其他债权投资信用减值准备'], ['5.现金流量套期储备'], ['6.外币财务报表折算差额'], ['7.其他'],
                    ['归属于少数股东的其他综合收益的税后净额'], ['七、综合收益总额'], ['归属于母公司所有者的综合收益总额'], ['归属于少数股东的综合收益总额'], ['八、每股收益：'],
                    ['（一）基本每股收益'], ['（二）稀释每股收益']]:
        st.append(_column)
        for _x in range(2):
            st[st.index(_column)].append('')
    num_table = gen_list('集团利润表', 6)
    match_1(num_table, st)
    match_2(num_table, '一、营业收入', st, '其中：营业收入')
    match_2(num_table, '减：营业成本', st, '其中：营业成本')
    match_2(num_table, '营业税金及附加', st, '税金及附加')
    match_2(num_table, '投资收益（损失以“-”号填列）', st, '投资收益（损失以“－”号填列）')
    match_2(num_table, '公允价值变动收益（损失以“-”号填列）', st, '公允价值变动收益（损失以“－”号填列）')
    match_2(num_table, '资产减值损失', st, '资产减值损失（损失以“-”号填列）', 0, 0, -1)
    match_3(st, '一、营业总收入', '=sum(B6:B9)')
    match_3(st, '二、营业总成本', '=sum(B11:B23)')
    match_3(st, '三、营业利润（亏损以“－”号填列）', '=B5-B10+sum(B26:B35)-B28-B29')
    match_3(st, '四、利润总额（亏损总额以“－”号填列）', '=B36+B37-B38')
    match_3(st, '五、净利润（净亏损以“－”号填列）', '=B39-B40')
    match_3(st, '一、营业总收入', '=sum(C6:C9)', 0, 2)
    match_3(st, '二、营业总成本', '=sum(C11:C23)', 0, 2)
    match_3(st, '三、营业利润（亏损以“－”号填列）', '=C5-C10+sum(C26:C35)-C28-C29', 0, 2)
    match_3(st, '四、利润总额（亏损总额以“－”号填列）', '=C36+C37-C38', 0, 2)
    match_3(st, '五、净利润（净亏损以“－”号填列）', '=C39-C40', 0, 2)
    _j = 0
    for _list in st:
        _i = -1
        for _value in _list:
            _i = _i + 1
            if isinstance(_value, str):
                _list[_i] = _value.replace('其它', '其他')
        st[_j] = _list
        _j = _j + 1
    # print('nt:', num_table)
    # print(st)
    return st


def format_preis():
    wb.sheets('preIS').range('A1:C80').api.Font.Size = 11
    wb.sheets('preIS').range('A1:A1').api.ColumnWidth = 55
    wb.sheets('preIS').range('B1:C1').api.ColumnWidth = 20
    wb.sheets('preIS').range('A:C').api.Font.Name = "Arial"
    wb.sheets('preIS').range('A1:C80').api.Style = "Comma"
    wb.sheets('preIS').range('A4:C70').api.Borders(7).LineStyle = 1
    wb.sheets('preIS').range('A4:C70').api.Borders(8).LineStyle = 1
    wb.sheets('preIS').range('A4:C70').api.Borders(9).LineStyle = 1
    wb.sheets('preIS').range('A4:C70').api.Borders(10).LineStyle = 1
    wb.sheets('preIS').range('A4:C70').api.Borders(11).LineStyle = 1
    wb.sheets('preIS').range('A4:C70').api.Borders(12).LineStyle = 1
    wb.sheets('preIS').range('A1:C1').api.HorizontalAlignment = 7
    wb.sheets('preIS').range('A5:C5').api.Font.Bold = True
    wb.sheets('preIS').range('A10:C10').api.Font.Bold = True
    wb.sheets('preIS').range('A36:C36').api.Font.Bold = True
    wb.sheets('preIS').range('A36:C36').api.Font.Bold = True
    wb.sheets('preIS').range('A39:C39').api.Font.Bold = True
    wb.sheets('preIS').range('A41:C41').api.Font.Bold = True
    wb.sheets('preIS').range('A48:C48').api.Font.Bold = True
    wb.sheets('preIS').range('A65:C65').api.Font.Bold = True
    wb.sheets('preIS').range('A68:C68').api.Font.Bold = True
    for lst in range(63):
        wb.sheets('preIS').range('a' + str(lst + 6)).api.IndentLevel = 1
    for lst in [10, 36, 39, 41, 48, 65, 68, 42, 45, 50, 56]:
        wb.sheets('preIS').range('a' + str(lst)).api.IndentLevel = 0
    for lst in [28, 29]:
        wb.sheets('preIS').range('a' + str(lst)).api.IndentLevel = 2
    return


# ----------------------------preIS Over-------------------------------------------
# ----------------------------preCF Start-------------------------------------------
def fill_precf():
    st = [['现金流量表', '', ''], ['编制单位:', '', '报表日期:'], [get_date_name()[1], '', get_date_name()[0]],
          ['项目:', '本期金额', '本年累计']]
    for _column in [['一、经营活动产生的现金流量：'], ['销售商品、提供劳务收到的现金'], ['客户存款和同业存放款项净增加额'], ['向中央银行借款净增加额'], ['向其他金融机构拆入资金净增加额'],
                    ['收到原保险合同保费取得的现金'], ['收到再保业务现金净额'], ['保户储金及投资款净增加额'], ['收取利息、手续费及佣金的现金'], ['拆入资金净增加额'],
                    ['回购业务资金净增加额'], ['代理买卖证券收到的现金净额'], ['收到的税费返还'], ['收到其它与经营活动有关的现金'], ['经营活动现金流入小计'],
                    ['购买商品、接受劳务支付的现金'], ['客户贷款及垫款净增加额'], ['存放中央银行和同业款项净增加额'], ['支付原保险合同赔付款项的现金'], ['拆出资金净增加额'],
                    ['支付利息、手续费及佣金的现金'], ['支付保单红利的现金'], ['支付给职工以及为职工支付的现金'], ['支付的各项税费'], ['支付其它与经营活动有关的现金'],
                    ['经营活动现金流出小计'], ['经营活动产生的现金流量净额'], ['二、投资活动产生的现金流量：'], ['收回投资收到的现金'], ['取得投资收益收到的现金'],
                    ['处置固定资产、无形资产和其它长期资产收回的现金净额'], ['处置子公司及其它营业单位收到的现金净额'], ['收到其它与投资活动有关的现金'], ['投资活动现金流入小计'],
                    ['购建固定资产、无形资产和其它长期资产支付的现金'], ['投资支付的现金'], ['质押贷款净增加额'], ['取得子公司及其它营业单位支付的现金净额'], ['支付其它与投资活动有关的现金'],
                    ['投资活动现金流出小计'], ['投资活动产生的现金流量净额'], ['三、筹资活动产生的现金流量：'], ['吸收投资收到的现金'], ['其中：子公司吸收少数股东投资收到的现金'],
                    ['取得借款收到的现金'], ['收到其它与筹资活动有关的现金'], ['筹资活动现金流入小计'], ['偿还债务支付的现金'], ['分配股利、利润或偿付利息支付的现金'],
                    ['支付其它与筹资活动有关的现金'], ['筹资活动现金流出小计'], ['筹资活动产生的现金流量净额'], ['四、汇率变动对现金及现金等价物的影响'], ['五、现金及现金等价物净增加额'],
                    ['加：期初现金及现金等价物余额'], ['六、期末现金及现金等价物余额']]:
        st.append(_column)
        for _x in range(2):
            st[st.index(_column)].append('')
        # print(st)
    num_table = gen_list('现金流量表', 6)
    match_1(num_table, st)
    match_3(st, '经营活动现金流入小计', '=sum(B6:B18)')
    match_3(st, '经营活动现金流出小计', '=sum(B20:B29)')
    match_3(st, '经营活动产生的现金流量净额', '=B19-B30')
    match_3(st, '投资活动现金流入小计', '=sum(B33:B37)')
    match_3(st, '投资活动现金流出小计', '=sum(B39:B43)')
    match_3(st, '投资活动产生的现金流量净额', '=B38-B44')
    match_3(st, '筹资活动现金流入小计', '=sum(B47,B49:B50)')
    match_3(st, '筹资活动现金流出小计', '=sum(B52:B54)')
    match_3(st, '筹资活动产生的现金流量净额', '=sum(B51-B55)')
    match_3(st, '五、现金及现金等价物净增加额', '=sum(B31,B45,B56,B57)')
    match_3(st, '六、期末现金及现金等价物余额', '=sum(B58,B59)')
    match_3(st, '经营活动现金流入小计', '=sum(C6:C18)', 0, 2)
    match_3(st, '经营活动现金流出小计', '=sum(C20:C29)', 0, 2)
    match_3(st, '经营活动产生的现金流量净额', '=C19-C30', 0, 2)
    match_3(st, '投资活动现金流入小计', '=sum(C33:C37)', 0, 2)
    match_3(st, '投资活动现金流出小计', '=sum(C39:C43)', 0, 2)
    match_3(st, '投资活动产生的现金流量净额', '=C38-C44', 0, 2)
    match_3(st, '筹资活动现金流入小计', '=sum(C47,C49:C50)', 0, 2)
    match_3(st, '筹资活动现金流出小计', '=sum(C52:C54)', 0, 2)
    match_3(st, '筹资活动产生的现金流量净额', '=sum(C51-C55)', 0, 2)
    match_3(st, '五、现金及现金等价物净增加额', '=sum(C31,C45,C56,C57)', 0, 2)
    match_3(st, '六、期末现金及现金等价物余额', '=sum(C58,C59)', 0, 2)
    _j = 0
    for _list in st:
        _i = -1
        for _value in _list:
            _i = _i + 1
            if isinstance(_value, str):
                _list[_i] = _value.replace('其它', '其他')
        st[_j] = _list
        _j = _j + 1
    # print(st)
    return st


def format_precf():
    wb.sheets('preCF').range('A1:C80').api.Font.Size = 11
    wb.sheets('preCF').range('A1:A80').api.ColumnWidth = 55
    wb.sheets('preCF').range('B1:C1').api.ColumnWidth = 25
    wb.sheets('preCF').range('A:C').api.Font.Name = "Arial"
    wb.sheets('preCF').range('A1:C80').api.Style = "Comma"
    wb.sheets('preCF').range('A4:C60').api.Borders(7).LineStyle = 1
    wb.sheets('preCF').range('A4:C60').api.Borders(8).LineStyle = 1
    wb.sheets('preCF').range('A4:C60').api.Borders(9).LineStyle = 1
    wb.sheets('preCF').range('A4:C60').api.Borders(10).LineStyle = 1
    wb.sheets('preCF').range('A4:C60').api.Borders(11).LineStyle = 1
    wb.sheets('preCF').range('A4:C60').api.Borders(12).LineStyle = 1
    wb.sheets('preCF').range('A1:C1').api.HorizontalAlignment = 7
    wb.sheets('preCF').range('A5:C5').api.Font.Bold = True
    wb.sheets('preCF').range('A19:C19').api.Font.Bold = True
    wb.sheets('preCF').range('A30:C32').api.Font.Bold = True
    wb.sheets('preCF').range('A38:C38').api.Font.Bold = True
    wb.sheets('preCF').range('A44:C46').api.Font.Bold = True
    wb.sheets('preCF').range('A51:C51').api.Font.Bold = True
    wb.sheets('preCF').range('A55:C58').api.Font.Bold = True
    wb.sheets('preCF').range('A60:C60').api.Font.Bold = True
    for lst in range(55):
        wb.sheets('preCF').range('a' + str(lst + 6)).api.IndentLevel = 1
    for lst in [19, 30, 31, 32, 38, 44, 45, 46, 51, 55, 56, 57, 58, 60]:
        wb.sheets('preCF').range('a' + str(lst)).api.IndentLevel = 0
    return


# ----------------------------preCF Over-------------------------------------------
# ----------------------------AJE Start-------------------------------------------
def fill_validation():
    wb.sheets('.Validation').range('A1').clear()
    wb.sheets('.Validation').range('A1').value = ['资产:'], ['流动资产：'], ['货币资金'], ['交易性金融资产'], ['衍生金融资产'], ['应收票据'], [
        '应收账款'], ['应收款项融资'], ['预付款项'], ['其他应收款'], ['其中：应收利息'], ['应收股利'], ['存货'], ['合同资产'], ['划分为持有待售的资产'], [
                                                     '一年内到期的非流动资产'], ['其他流动资产'], ['流动资产合计'], ['非流动资产：'], ['发放贷款和垫款'], [
                                                     '债权投资'], ['其他债权投资'], ['长期应收款'], ['长期股权投资'], ['其他权益工具投资'], [
                                                     '其他非流动金融资产'], ['投资性房地产'], ['固定资产'], ['在建工程'], ['生产性生物资产'], [
                                                     '油气资产'], ['使用权资产'], ['无形资产'], ['开发支出'], ['商誉'], ['长期待摊费用'], [
                                                     '递延所得税资产'], ['其他非流动资产'], ['非流动资产合计'], ['资产总计'], ['负债和股东权益:'], [
                                                     '流动负债：'], ['短期借款'], ['交易性金融负债'], ['衍生金融负债'], ['应付票据'], ['应付账款'], [
                                                     '预收款项'], ['合同负债'], ['应付职工薪酬'], ['应交税费'], ['其他应付款'], ['其中：应付利息'], [
                                                     '应付股利'], ['划分为持有待售的负债'], ['一年内到期的非流动负债'], ['其他流动负债'], ['流动负债合计'], [
                                                     '非流动负债：'], ['保险合同准备金'], ['长期借款'], ['应付债券'], ['其中：优先股(负债)'], [
                                                     '永续债(负债)'], ['租赁负债'], ['长期应付款'], ['长期应付职工薪酬'], ['预计负债'], [
                                                     '递延收益'], ['递延所得税负债'], ['其他非流动负债'], ['非流动负债合计'], ['负债合计'], [
                                                     '所有者权益：'], ['实收资本(股本)'], ['其他权益工具'], ['其中：优先股(权益)'], ['永续债(权益)'], [
                                                     '资本公积'], ['减：库存股'], ['其他综合收益'], ['专项储备'], ['盈余公积'], ['一般风险准备'], [
                                                     '未分配利润'], ['归属于母公司所有者权益合计'], ['少数股东权益'], ['股东权益合计'], [
                                                     '负债和股东权益总计'], ['一、营业总收入'], ['其中：营业收入'], ['利息收入_'], ['已赚保费'], [
                                                     '手续费及佣金收入'], ['二、营业总成本'], ['其中：营业成本'], ['利息支出'], ['手续费及佣金支出'], [
                                                     '退保金'], ['赔付支出净额'], ['提取保险责任准备金净额'], ['保单红利支出'], ['分保费用'], [
                                                     '税金及附加'], ['销售费用'], ['管理费用'], ['研发费用'], ['财务费用'], ['其中：利息费用'], [
                                                     '利息收入'], ['加：其他收益'], ['投资收益（损失以“－”号填列）'], ['其中：对联营企业和合营企业的投资收益'], [
                                                     '以摊余成本计量的金融资产终止确认收益'], ['汇兑收益（损失以“-”号填列）'], [
                                                     '净敞口套期收益（损失以“－”号填列）'], ['公允价值变动收益（损失以“－”号填列）'], [
                                                     '信用减值损失（损失以“-”号填列）'], ['资产减值损失（损失以“-”号填列）'], [
                                                     '资产处置收益（损失以“-”号填列）'], ['三、营业利润（亏损以“－”号填列）'], ['加：营业外收入'], [
                                                     '减：营业外支出'], ['四、利润总额（亏损总额以“－”号填列）'], ['减：所得税费用'], [
                                                     '五、净利润（净亏损以“－”号填列）'], ['（一）按经营持续性分类'], ['1.持续经营净利润（净亏损以“－”号填列）'], [
                                                     '2.终止经营净利润（净亏损以“－”号填列）'], ['（二）按所有权归属分类'], ['1.归属于母公司所有者的净利润'], [
                                                     '2.少数股东损益'], ['六、其他综合收益的税后净额'], ['归属母公司所有者的其他综合收益的税后净额'], [
                                                     '（一）不能重分类进损益的其他综合收益'], ['1.重新计量设定受益计划变动额'], [
                                                     '2.权益法下不能转损益的其他综合收益'], ['3.其他权益工具投资公允价值变动'], [
                                                     '4.企业自身信用风险公允价值变动'], ['5.其他'], ['（二）将重分类进损益的其他综合收益'], [
                                                     '1.权益法下可转损益的其他综合收益'], ['2.其他债权投资公允价值变动'], [
                                                     '3.金融资产重分类计入其他综合收益的金额'], ['4.其他债权投资信用减值准备'], ['5.现金流量套期储备'], [
                                                     '6.外币财务报表折算差额'], ['7.其他'], ['归属于少数股东的其他综合收益的税后净额'], ['七、综合收益总额'], [
                                                     '归属于母公司所有者的综合收益总额'], ['归属于少数股东的综合收益总额'], ['八、每股收益：'], [
                                                     '（一）基本每股收益'], ['（二）稀释每股收益'], ['一、经营活动产生的现金流量：'], [
                                                     '销售商品、提供劳务收到的现金'], ['客户存款和同业存放款项净增加额'], ['向中央银行借款净增加额'], [
                                                     '向其他金融机构拆入资金净增加额'], ['收到原保险合同保费取得的现金'], ['收到再保业务现金净额'], [
                                                     '保户储金及投资款净增加额'], ['收取利息、手续费及佣金的现金'], ['拆入资金净增加额'], [
                                                     '回购业务资金净增加额'], ['代理买卖证券收到的现金净额'], ['收到的税费返还'], [
                                                     '收到其他与经营活动有关的现金'], ['经营活动现金流入小计'], ['购买商品、接受劳务支付的现金'], [
                                                     '客户贷款及垫款净增加额'], ['存放中央银行和同业款项净增加额'], ['支付原保险合同赔付款项的现金'], [
                                                     '拆出资金净增加额'], ['支付利息、手续费及佣金的现金'], ['支付保单红利的现金'], [
                                                     '支付给职工以及为职工支付的现金'], ['支付的各项税费'], ['支付其他与经营活动有关的现金'], [
                                                     '经营活动现金流出小计'], ['经营活动产生的现金流量净额'], ['二、投资活动产生的现金流量：'], [
                                                     '收回投资收到的现金'], ['取得投资收益收到的现金'], ['处置固定资产、无形资产和其他长期资产收回的现金净额'], [
                                                     '处置子公司及其他营业单位收到的现金净额'], ['收到其他与投资活动有关的现金'], ['投资活动现金流入小计'], [
                                                     '购建固定资产、无形资产和其他长期资产支付的现金'], ['投资支付的现金'], ['质押贷款净增加额'], [
                                                     '取得子公司及其他营业单位支付的现金净额'], ['支付其他与投资活动有关的现金'], ['投资活动现金流出小计'], [
                                                     '投资活动产生的现金流量净额'], ['三、筹资活动产生的现金流量：'], ['吸收投资收到的现金'], [
                                                     '其中：子公司吸收少数股东投资收到的现金'], ['取得借款收到的现金'], ['收到其他与筹资活动有关的现金'], [
                                                     '筹资活动现金流入小计'], ['偿还债务支付的现金'], ['分配股利、利润或偿付利息支付的现金'], [
                                                     '支付其他与筹资活动有关的现金'], ['筹资活动现金流出小计'], ['筹资活动产生的现金流量净额'], [
                                                     '四、汇率变动对现金及现金等价物的影响'], ['五、现金及现金等价物净增加额'], ['加：期初现金及现金等价物余额'], [
                                                     '六、期末现金及现金等价物余额']
    return


def set_validation(range_1):
    range_1.api.Validation.Delete()
    range_1.api.Validation.Add(3, 3, 1, '=\'.Validation\'!$A$1:$A$220')
    return


#
def fill_aje():
    wb.sheets('AJE').clear()
    st = [['摘要', '借贷方向', '一级科目', '二级科目', '借方金额', '贷方金额', '备注']]
    # 提示科目是否写对
    for _row in range(500):
        st.append(['', '', '', '', '', '', '=IF(OR(NOT(ISERROR(VLOOKUP(C' + str(_row + 2) \
                   + ',\'.Validation\'!A:A,1,0))),C' + str(_row + 2) + '=""),"","没有找到该科目,请检查")'])
    # 填写固定调整分录
    if plookup('preBS', '应交税费', 4) < 0:
        rplc_lst(st, 1, ['将应交税费负数调整到其他流动资产', '借', '其他流动资产', '', -plookup('preBS', '应交税费', 4), '', ''])
        rplc_lst(st, 1, ['将应交税费负数调整到其他流动资产', '贷', '应交税费', '', '', -plookup('preBS', '应交税费', 4), ''])
        rplc_lst(st, 1, ['-', '', '', '', '', '', ''])
    # print(st)
    rplc_lst(st, 1, ['将工程设备款重分类至其他非流动资产核算', '借', '其他非流动资产', '', '', '', ''])
    rplc_lst(st, 1, ['将工程设备款重分类至其他非流动资产核算', '贷', '预付款项', '', '', '', ''])
    rplc_lst(st, 1, ['-', '', '', '', '', '', ''])
    return st


def plookup(sheet_name, head, col_1, col_2=1):
    lkupval = ''
    _num_table = wb.sheets(sheet_name).range("A1:F80").value
    for _lst in _num_table:
        if _lst[col_1 - 1] == head:
            lkupval = _lst[col_1 - 1 + col_2]
            break
        else:
            lkupval = 0
    return lkupval


def rplc_lst(st, col, lst):
    _i = 0
    for _lst in st:
        if _lst[col - 1] == '' and _lst[col] == '':
            st[_i][:6] = lst[:6]
            # st[st.index(_lst)][0] = '-'
            break
        _i = _i + 1
    return


def format_aje():
    wb.sheets('AJE').range('A:G').api.Font.Size = 11
    wb.sheets('AJE').range('A1').api.ColumnWidth = 55
    wb.sheets('AJE').range('C1:G1').api.ColumnWidth = 20
    wb.sheets('AJE').range('E:F').api.Style = "Comma"
    return


# ----------------------------AJE Over-------------------------------------------
# ----------------------------TB Start-------------------------------------------
def fill_tb():
    st = [['科目名称', '调整前', '借方调整数', '贷方调整数', '调整后']]
    for _lst in wb.sheets('preBS').range('A4:A52').value:
        if _lst is not None:
            # print(_lst)
            st.append([_lst, '', '', '', ''])
    # print(st)
    for _lst in wb.sheets('preBS').range('D4:D52').value:
        st.append([_lst, '', '', '', ''])
    # print(st)
    st.append(['', '', '', '', ''])
    st.append(['利润表项目:', '', '', '', ''])
    for _lst in wb.sheets('preIS').range('A5:A70').value:
        st.append([_lst, '', '', '', ''])
    st.append(['', '', '', '', ''])
    st.append(['现流表项目:', '', '', '', ''])
    for _lst in wb.sheets('preCF').range('A5:A60').value:
        st.append([_lst, '', '', '', ''])

    _i = 1
    for _lst in st[1:]:
        _i = _i + 1
        st[st.index(_lst)][1] = '=IFERROR(IFERROR(IFERROR(VLOOKUP(A' + str(_i) + ',preBS!A:B,2,0),VLOOKUP(A' + str(
            _i) + ',preBS!D:E,2,0)),VLOOKUP(A' + str(_i) + ',preIS!A:B,2,0)),VLOOKUP(A' + str(_i) + ',preCF!A:B,2,0)) '
        st[st.index(_lst)][2] = '=IF(SUMIF(AJE!C:C,A' + str(_i) + ',AJE!E:E)=0,0,SUMIF(AJE!C:C,A' + str(
            _i) + ',AJE!E:E)) '
        st[st.index(_lst)][3] = '=IF(SUMIF(AJE!C:C,A' + str(_i) + ',AJE!F:F)=0,0,SUMIF(AJE!C:C,A' + str(
            _i) + ',AJE!F:F)) '

    #     区分借贷方向后编制'调整后'列公式
    dict_dc = {'资产:': 1, '流动资产：': 1, '货币资金': 1, '交易性金融资产': 1, '衍生金融资产': 1, '应收票据': 1, '应收账款': 1, '应收款项融资': 1, '预付款项': 1,
               '其他应收款': 1, '其中：应收利息': 1, '应收股利': 1, '存货': 1, '合同资产': 1, '划分为持有待售的资产': 1, '一年内到期的非流动资产': 1, '其他流动资产': 1,
               '流动资产合计': 1, '非流动资产：': 1, '发放贷款和垫款': 1, '债权投资': 1, '其他债权投资': 1, '长期应收款': 1, '长期股权投资': 1, '其他权益工具投资': 1,
               '其他非流动金融资产': 1, '投资性房地产': 1, '固定资产': 1, '在建工程': 1, '生产性生物资产': 1, '油气资产': 1, '使用权资产': 1, '无形资产': 1,
               '开发支出': 1, '商誉': 1, '长期待摊费用': 1, '递延所得税资产': 1, '其他非流动资产': 1, '非流动资产合计': 1, '资产总计': 1, '负债和股东权益:': -1,
               '流动负债：': -1, '短期借款': -1, '交易性金融负债': -1, '衍生金融负债': -1, '应付票据': -1, '应付账款': -1, '预收款项': -1, '合同负债': -1,
               '应付职工薪酬': -1, '应交税费': -1, '其他应付款': -1, '其中：应付利息': -1, '应付股利': -1, '划分为持有待售的负债': -1, '一年内到期的非流动负债': -1,
               '其他流动负债': -1, '流动负债合计': -1, '非流动负债：': -1, '保险合同准备金': -1, '长期借款': -1, '应付债券': -1, '其中：优先股(负债)': -1,
               '永续债(负债)': -1, '租赁负债': -1, '长期应付款': -1, '长期应付职工薪酬': -1, '预计负债': -1, '递延收益': -1, '递延所得税负债': -1,
               '其他非流动负债': -1, '非流动负债合计': -1, '负债合计': -1, '所有者权益：': -1, '实收资本(股本)': -1, '其他权益工具': -1, '其中：优先股': -1,
               '永续债': -1, '资本公积': -1, '减：库存股': -1, '其他综合收益': -1, '专项储备': -1, '盈余公积': -1, '一般风险准备': -1, '未分配利润': -1,
               '归属于母公司所有者权益合计': -1, '少数股东权益': -1, '股东权益合计': -1, '负债和股东权益总计': -1, '利润表项目:': -1, '一、营业总收入': -1,
               '其中：营业收入': -1, '利息收入_': -1, '已赚保费': -1, '手续费及佣金收入': -1, '二、营业总成本': 1, '其中：营业成本': 1, '利息支出': 1,
               '手续费及佣金支出': 1, '退保金': 1, '赔付支出净额': 1, '提取保险责任准备金净额': 1, '保单红利支出': 1, '分保费用': 1, '税金及附加': 1, '销售费用': 1,
               '管理费用': 1, '研发费用': 1, '财务费用': 1, '其中：利息费用': 1, '利息收入': -1, '加：其他收益': -1, '投资收益（损失以“－”号填列）': -1,
               '其中：对联营企业和合营企业的投资收益': -1, '以摊余成本计量的金融资产终止确认收益': -1, '汇兑收益（损失以“-”号填列）': -1, '净敞口套期收益（损失以“－”号填列）': -1,
               '公允价值变动收益（损失以“－”号填列）': -1, '信用减值损失（损失以“-”号填列）': -1, '资产减值损失（损失以“-”号填列）': -1, '资产处置收益（损失以“-”号填列）': -1,
               '三、营业利润（亏损以“－”号填列）': -1, '加：营业外收入': -1, '减：营业外支出': 1, '四、利润总额（亏损总额以“－”号填列）': -1, '减：所得税费用': 1,
               '五、净利润（净亏损以“－”号填列）': -1, '（一）按经营持续性分类': -1, '1.持续经营净利润（净亏损以“－”号填列）': 1, '2.终止经营净利润（净亏损以“－”号填列）': 1,
               '（二）按所有权归属分类': 1, '1.归属于母公司所有者的净利润': 1, '2.少数股东损益': 1, '六、其他综合收益的税后净额': -1, '归属母公司所有者的其他综合收益的税后净额': -1,
               '（一）不能重分类进损益的其他综合收益': -1, '1.重新计量设定受益计划变动额': -1, '2.权益法下不能转损益的其他综合收益': -1, '3.其他权益工具投资公允价值变动': -1,
               '4.企业自身信用风险公允价值变动': -1, '5.其他': -1, '（二）将重分类进损益的其他综合收益': -1, '1.权益法下可转损益的其他综合收益': -1,
               '2.其他债权投资公允价值变动': -1, '3.金融资产重分类计入其他综合收益的金额': -1, '4.其他债权投资信用减值准备': -1, '5.现金流量套期储备': -1,
               '6.外币财务报表折算差额': -1, '7.其他': -1, '归属于少数股东的其他综合收益的税后净额': -1, '七、综合收益总额': -1, '归属于母公司所有者的综合收益总额': -1,
               '归属于少数股东的综合收益总额': 1, '八、每股收益：': 1, '（一）基本每股收益': 1, '（二）稀释每股收益': 1, '现流表项目:': 1, '一、经营活动产生的现金流量：': 1,
               '销售商品、提供劳务收到的现金': 1, '客户存款和同业存放款项净增加额': 1, '向中央银行借款净增加额': 1, '向其他金融机构拆入资金净增加额': 1, '收到原保险合同保费取得的现金': 1,
               '收到再保业务现金净额': 1, '保户储金及投资款净增加额': 1, '收取利息、手续费及佣金的现金': 1, '拆入资金净增加额': 1, '回购业务资金净增加额': 1,
               '代理买卖证券收到的现金净额': 1, '收到的税费返还': 1, '收到其他与经营活动有关的现金': 1, '经营活动现金流入小计': 1, '购买商品、接受劳务支付的现金': 1,
               '客户贷款及垫款净增加额': 1, '存放中央银行和同业款项净增加额': 1, '支付原保险合同赔付款项的现金': 1, '拆出资金净增加额': 1, '支付利息、手续费及佣金的现金': 1,
               '支付保单红利的现金': 1, '支付给职工以及为职工支付的现金': 1, '支付的各项税费': 1, '支付其他与经营活动有关的现金': 1, '经营活动现金流出小计': 1,
               '经营活动产生的现金流量净额': 1, '二、投资活动产生的现金流量：': 1, '收回投资收到的现金': 1, '取得投资收益收到的现金': 1,
               '处置固定资产、无形资产和其他长期资产收回的现金净额': 1, '处置子公司及其他营业单位收到的现金净额': 1, '收到其他与投资活动有关的现金': 1, '投资活动现金流入小计': 1,
               '购建固定资产、无形资产和其他长期资产支付的现金': 1, '投资支付的现金': 1, '质押贷款净增加额': 1, '取得子公司及其他营业单位支付的现金净额': 1, '支付其他与投资活动有关的现金': 1,
               '投资活动现金流出小计': 1, '投资活动产生的现金流量净额': 1, '三、筹资活动产生的现金流量：': 1, '吸收投资收到的现金': 1, '其中：子公司吸收少数股东投资收到的现金': 1,
               '取得借款收到的现金': 1, '收到其他与筹资活动有关的现金': 1, '筹资活动现金流入小计': 1, '偿还债务支付的现金': 1, '分配股利、利润或偿付利息支付的现金': 1,
               '支付其他与筹资活动有关的现金': 1, '筹资活动现金流出小计': 1, '筹资活动产生的现金流量净额': 1, '四、汇率变动对现金及现金等价物的影响': 1, '五、现金及现金等价物净增加额': 1,
               '加：期初现金及现金等价物余额': 1, '六、期末现金及现金等价物余额': 1
               }
    _i = 0
    for _lst in st:
        _i = _i + 1
        if _lst[0] in dict_dc:
            if dict_dc[_lst[0]] == 1:
                st[st.index(_lst)][4] = '=B' + str(_i) + '+ IF(C' + str(_i) + '="", 0, C' + str(_i) + ') - IF(D' + str(
                    _i) + '="", 0, D' + str(_i) + ')'
            elif dict_dc[_lst[0]] == -1:
                st[st.index(_lst)][4] = '=B' + str(_i) + '- IF(C' + str(_i) + '="", 0, C' + str(_i) + ') + IF(D' + str(
                    _i) + '="", 0, D' + str(_i) + ')'

    # 修正无需计算的行
    for _lst in st:
        if _lst[0] == '' or _lst[0][-1] == ':' or _lst[0][-1] == '：':
            st[st.index(_lst)][1] = st[st.index(_lst)][2] = st[st.index(_lst)][3] = st[st.index(_lst)][4] = ''
    # 修正用公式计算的行
    _st_0 = ['zero']
    for _lst in st:
        _st_0.append(_lst[0])
    for _lst in st:
        # print(_st_0)
        # print(_lst[0])
        if _lst[0] == '流动资产合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('货币资金')) + ':B' + str(
                _st_0.index('其他流动资产')) + ')-B' + str(_st_0.index('其中：应收利息')) + '-B' + str(_st_0.index('应收股利'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('货币资金')) + ':C' + str(
                _st_0.index('其他流动资产')) + ')-C' + str(_st_0.index('其中：应收利息')) + '-C' + str(_st_0.index('应收股利'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('货币资金')) + ':D' + str(
                _st_0.index('其他流动资产')) + ')-D' + str(_st_0.index('其中：应收利息')) + '-D' + str(_st_0.index('应收股利'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('货币资金')) + ':E' + str(
                _st_0.index('其他流动资产')) + ')-E' + str(_st_0.index('其中：应收利息')) + '-E' + str(_st_0.index('应收股利'))
        if _lst[0] == '非流动资产合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('发放贷款和垫款')) + ':B' + str(_st_0.index('其他非流动资产'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('发放贷款和垫款')) + ':C' + str(_st_0.index('其他非流动资产'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('发放贷款和垫款')) + ':D' + str(_st_0.index('其他非流动资产'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('发放贷款和垫款')) + ':E' + str(_st_0.index('其他非流动资产'))
        if _lst[0] == '资产总计':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('流动资产合计')) + '+B' + str(_st_0.index('非流动资产合计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('流动资产合计')) + '+C' + str(_st_0.index('非流动资产合计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('流动资产合计')) + '+D' + str(_st_0.index('非流动资产合计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('流动资产合计')) + '+E' + str(_st_0.index('非流动资产合计'))
        if _lst[0] == '流动负债合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('短期借款')) + ':B' + str(
                _st_0.index('其他流动负债')) + ')-B' + str(_st_0.index('其中：应付利息')) + '-B' + str(_st_0.index('应付股利'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('短期借款')) + ':C' + str(
                _st_0.index('其他流动负债')) + ')-C' + str(_st_0.index('其中：应付利息')) + '-C' + str(_st_0.index('应付股利'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('短期借款')) + ':D' + str(
                _st_0.index('其他流动负债')) + ')-D' + str(_st_0.index('其中：应付利息')) + '-D' + str(_st_0.index('应付股利'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('短期借款')) + ':E' + str(
                _st_0.index('其他流动负债')) + ')-E' + str(_st_0.index('其中：应付利息')) + '-E' + str(_st_0.index('应付股利'))
        if _lst[0] == '非流动负债合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('保险合同准备金')) + ':B' + str(
                _st_0.index('其他非流动负债')) + ')-B' + str(
                _st_0.index('其中：优先股(负债)')) + '-B' + str(_st_0.index('永续债(负债)'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('保险合同准备金')) + ':C' + str(
                _st_0.index('其他非流动负债')) + ')-C' + str(
                _st_0.index('其中：优先股(负债)')) + '-C' + str(_st_0.index('永续债(负债)'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('保险合同准备金')) + ':D' + str(
                _st_0.index('其他非流动负债')) + ')-D' + str(
                _st_0.index('其中：优先股(负债)')) + '-D' + str(_st_0.index('永续债(负债)'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('保险合同准备金')) + ':E' + str(
                _st_0.index('其他非流动负债')) + ')-E' + str(
                _st_0.index('其中：优先股(负债)')) + '-E' + str(_st_0.index('永续债(负债)'))
        if _lst[0] == '负债合计':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('流动负债合计')) + '+B' + str(_st_0.index('非流动负债合计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('流动负债合计')) + '+C' + str(_st_0.index('非流动负债合计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('流动负债合计')) + '+D' + str(_st_0.index('非流动负债合计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('流动负债合计')) + '+E' + str(_st_0.index('非流动负债合计'))
        if _lst[0] == '归属于母公司所有者权益合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('实收资本(股本)')) + ':B' + str(
                _st_0.index('未分配利润')) + ',-B' + str(_st_0.index('其中：优先股(权益)')) + ',-B' + str(
                _st_0.index('永续债(权益)')) + ',-2*B' + str(_st_0.index('减：库存股'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('实收资本(股本)')) + ':C' + str(
                _st_0.index('未分配利润')) + ',-C' + str(_st_0.index('其中：优先股(权益)')) + ',-C' + str(
                _st_0.index('永续债(权益)')) + ',-2*C' + str(_st_0.index('减：库存股'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('实收资本(股本)')) + ':D' + str(
                _st_0.index('未分配利润')) + ',-D' + str(_st_0.index('其中：优先股(权益)')) + ',-D' + str(
                _st_0.index('永续债(权益)')) + ',-2*D' + str(_st_0.index('减：库存股'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('实收资本(股本)')) + ':E' + str(
                _st_0.index('未分配利润')) + ',-E' + str(_st_0.index('其中：优先股(权益)')) + ',-E' + str(
                _st_0.index('永续债(权益)')) + ',-2*E' + str(_st_0.index('减：库存股'))
        if _lst[0] == '股东权益合计':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('归属于母公司所有者权益合计')) + '+B' + str(_st_0.index('少数股东权益'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('归属于母公司所有者权益合计')) + '+C' + str(_st_0.index('少数股东权益'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('归属于母公司所有者权益合计')) + '+D' + str(_st_0.index('少数股东权益'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('归属于母公司所有者权益合计')) + '+E' + str(_st_0.index('少数股东权益'))
        if _lst[0] == '负债和股东权益总计':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('负债合计')) + '+B' + str(_st_0.index('股东权益合计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('负债合计')) + '+C' + str(_st_0.index('股东权益合计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('负债合计')) + '+D' + str(_st_0.index('股东权益合计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('负债合计')) + '+E' + str(_st_0.index('股东权益合计'))
        if _lst[0] == '一、营业总收入':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('其中：营业收入')) + ':B' + str(_st_0.index('手续费及佣金收入'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('其中：营业收入')) + ':C' + str(_st_0.index('手续费及佣金收入'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('其中：营业收入')) + ':D' + str(_st_0.index('手续费及佣金收入'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('其中：营业收入')) + ':E' + str(_st_0.index('手续费及佣金收入'))
        if _lst[0] == '二、营业总成本':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('其中：营业成本')) + ':B' + str(_st_0.index('财务费用'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('其中：营业成本')) + ':C' + str(_st_0.index('财务费用'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('其中：营业成本')) + ':D' + str(_st_0.index('财务费用'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('其中：营业成本')) + ':E' + str(_st_0.index('财务费用'))
        if _lst[0] == '三、营业利润（亏损以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('一、营业总收入')) + '-B' + str(
                _st_0.index('二、营业总成本')) + '+SUM(B' + str(_st_0.index('加：其他收益')) + ':B' + str(
                _st_0.index('资产处置收益（损失以“-”号填列）')) + ')-B' + str(_st_0.index('其中：对联营企业和合营企业的投资收益')) + '-B' + str(
                _st_0.index('以摊余成本计量的金融资产终止确认收益'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('一、营业总收入')) + '-C' + str(
                _st_0.index('二、营业总成本')) + '+SUM(C' + str(_st_0.index('加：其他收益')) + ':C' + str(
                _st_0.index('资产处置收益（损失以“-”号填列）')) + ')-C' + str(_st_0.index('其中：对联营企业和合营企业的投资收益')) + '-C' + str(
                _st_0.index('以摊余成本计量的金融资产终止确认收益'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('一、营业总收入')) + '-D' + str(
                _st_0.index('二、营业总成本')) + '+SUM(D' + str(_st_0.index('加：其他收益')) + ':D' + str(
                _st_0.index('资产处置收益（损失以“-”号填列）')) + ')-D' + str(_st_0.index('其中：对联营企业和合营企业的投资收益')) + '-D' + str(
                _st_0.index('以摊余成本计量的金融资产终止确认收益'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('一、营业总收入')) + '-E' + str(
                _st_0.index('二、营业总成本')) + '+SUM(E' + str(_st_0.index('加：其他收益')) + ':E' + str(
                _st_0.index('资产处置收益（损失以“-”号填列）')) + ')-E' + str(_st_0.index('其中：对联营企业和合营企业的投资收益')) + '-E' + str(
                _st_0.index('以摊余成本计量的金融资产终止确认收益'))
        if _lst[0] == '四、利润总额（亏损总额以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('三、营业利润（亏损以“－”号填列）')) + '+B' + str(
                _st_0.index('加：营业外收入')) + '-B' + str(_st_0.index('减：营业外支出'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('三、营业利润（亏损以“－”号填列）')) + '+C' + str(
                _st_0.index('加：营业外收入')) + '-C' + str(_st_0.index('减：营业外支出'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('三、营业利润（亏损以“－”号填列）')) + '+D' + str(
                _st_0.index('加：营业外收入')) + '-D' + str(_st_0.index('减：营业外支出'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('三、营业利润（亏损以“－”号填列）')) + '+E' + str(
                _st_0.index('加：营业外收入')) + '-E' + str(_st_0.index('减：营业外支出'))
        if _lst[0] == '五、净利润（净亏损以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('四、利润总额（亏损总额以“－”号填列）')) + '-B' + str(
                _st_0.index('减：所得税费用'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('四、利润总额（亏损总额以“－”号填列）')) + '-C' + str(
                _st_0.index('减：所得税费用'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('四、利润总额（亏损总额以“－”号填列）')) + '-D' + str(
                _st_0.index('减：所得税费用'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('四、利润总额（亏损总额以“－”号填列）')) + '-E' + str(
                _st_0.index('减：所得税费用'))
        if _lst[0] == '经营活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('销售商品、提供劳务收到的现金')) + ':B' + str(
                _st_0.index('收到其他与经营活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('销售商品、提供劳务收到的现金')) + ':C' + str(
                _st_0.index('收到其他与经营活动有关的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('销售商品、提供劳务收到的现金')) + ':D' + str(
                _st_0.index('收到其他与经营活动有关的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('销售商品、提供劳务收到的现金')) + ':E' + str(
                _st_0.index('收到其他与经营活动有关的现金'))
        if _lst[0] == '经营活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('购买商品、接受劳务支付的现金')) + ':B' + str(
                _st_0.index('支付其他与经营活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('购买商品、接受劳务支付的现金')) + ':C' + str(
                _st_0.index('支付其他与经营活动有关的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('购买商品、接受劳务支付的现金')) + ':D' + str(
                _st_0.index('支付其他与经营活动有关的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('购买商品、接受劳务支付的现金')) + ':E' + str(
                _st_0.index('支付其他与经营活动有关的现金'))
        if _lst[0] == '经营活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('经营活动现金流入小计')) + '-B' + str(
                _st_0.index('经营活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('经营活动现金流入小计')) + '-C' + str(
                _st_0.index('经营活动现金流出小计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('经营活动现金流入小计')) + '-D' + str(
                _st_0.index('经营活动现金流出小计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('经营活动现金流入小计')) + '-E' + str(
                _st_0.index('经营活动现金流出小计'))
        if _lst[0] == '投资活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('收回投资收到的现金')) + ':B' + str(
                _st_0.index('收到其他与投资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('收回投资收到的现金')) + ':C' + str(
                _st_0.index('收到其他与投资活动有关的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('收回投资收到的现金')) + ':D' + str(
                _st_0.index('收到其他与投资活动有关的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('收回投资收到的现金')) + ':E' + str(
                _st_0.index('收到其他与投资活动有关的现金'))
        if _lst[0] == '投资活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':B' + str(
                _st_0.index('支付其他与投资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':C' + str(
                _st_0.index('支付其他与投资活动有关的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':D' + str(
                _st_0.index('支付其他与投资活动有关的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':E' + str(
                _st_0.index('支付其他与投资活动有关的现金'))
        if _lst[0] == '投资活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('投资活动现金流入小计')) + '-B' + str(
                _st_0.index('投资活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('投资活动现金流入小计')) + '-C' + str(
                _st_0.index('投资活动现金流出小计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('投资活动现金流入小计')) + '-D' + str(
                _st_0.index('投资活动现金流出小计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('投资活动现金流入小计')) + '-E' + str(
                _st_0.index('投资活动现金流出小计'))
        if _lst[0] == '筹资活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('取得借款收到的现金')) + ':B' + str(
                _st_0.index('收到其他与筹资活动有关的现金')) + ',B' + str(_st_0.index('吸收投资收到的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('取得借款收到的现金')) + ':C' + str(
                _st_0.index('收到其他与筹资活动有关的现金')) + ',C' + str(_st_0.index('吸收投资收到的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('取得借款收到的现金')) + ':D' + str(
                _st_0.index('收到其他与筹资活动有关的现金')) + ',D' + str(_st_0.index('吸收投资收到的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('取得借款收到的现金')) + ':E' + str(
                _st_0.index('收到其他与筹资活动有关的现金')) + ',E' + str(_st_0.index('吸收投资收到的现金'))
        if _lst[0] == '筹资活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('偿还债务支付的现金')) + ':B' + str(
                _st_0.index('支付其他与筹资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('偿还债务支付的现金')) + ':C' + str(
                _st_0.index('支付其他与筹资活动有关的现金'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('偿还债务支付的现金')) + ':D' + str(
                _st_0.index('支付其他与筹资活动有关的现金'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('偿还债务支付的现金')) + ':E' + str(
                _st_0.index('支付其他与筹资活动有关的现金'))
        if _lst[0] == '筹资活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('筹资活动现金流入小计')) + '-B' + str(
                _st_0.index('筹资活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('筹资活动现金流入小计')) + '-C' + str(
                _st_0.index('筹资活动现金流出小计'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('筹资活动现金流入小计')) + '-D' + str(
                _st_0.index('筹资活动现金流出小计'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('筹资活动现金流入小计')) + '-E' + str(
                _st_0.index('筹资活动现金流出小计'))
        if _lst[0] == '五、现金及现金等价物净增加额':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('经营活动产生的现金流量净额')) + ',B' + str(
                _st_0.index('投资活动产生的现金流量净额')) + ',B' + str(_st_0.index('筹资活动产生的现金流量净额')) + ',B' + str(
                _st_0.index('四、汇率变动对现金及现金等价物的影响'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('经营活动产生的现金流量净额')) + ',C' + str(
                _st_0.index('投资活动产生的现金流量净额')) + ',C' + str(_st_0.index('筹资活动产生的现金流量净额')) + ',C' + str(
                _st_0.index('四、汇率变动对现金及现金等价物的影响'))
            st[st.index(_lst)][3] = '=SUM(D' + str(_st_0.index('经营活动产生的现金流量净额')) + ',D' + str(
                _st_0.index('投资活动产生的现金流量净额')) + ',D' + str(_st_0.index('筹资活动产生的现金流量净额')) + ',D' + str(
                _st_0.index('四、汇率变动对现金及现金等价物的影响'))
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_0.index('经营活动产生的现金流量净额')) + ',E' + str(
                _st_0.index('投资活动产生的现金流量净额')) + ',E' + str(_st_0.index('筹资活动产生的现金流量净额')) + ',E' + str(
                _st_0.index('四、汇率变动对现金及现金等价物的影响'))
        if _lst[0] == '六、期末现金及现金等价物余额':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('五、现金及现金等价物净增加额')) + '+B' + str(
                _st_0.index('加：期初现金及现金等价物余额'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('五、现金及现金等价物净增加额')) + '+C' + str(
                _st_0.index('加：期初现金及现金等价物余额'))
            st[st.index(_lst)][3] = '=D' + str(_st_0.index('五、现金及现金等价物净增加额')) + '+D' + str(
                _st_0.index('加：期初现金及现金等价物余额'))
            st[st.index(_lst)][4] = '=E' + str(_st_0.index('五、现金及现金等价物净增加额')) + '+E' + str(
                _st_0.index('加：期初现金及现金等价物余额'))
    return st


def format_tb():
    wb.sheets('TB').range('B:E').api.NumberFormatLocal = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ""-""??_ ;_ @_ "
    wb.sheets('TB').range('A:E').api.Font.Size = 11
    wb.sheets('TB').range('A1').api.ColumnWidth = 50
    wb.sheets('TB').range('B:E').api.ColumnWidth = 30
    return


# ----------------------------TB Over-------------------------------------------
# ----------------------------BS Start-------------------------------------------
def fill_bs():
    st = wb.sheets('preBS').range('A1:F52').value
    # 修正无需计算的行
    for _lst in st[5:]:
        if _lst[0] is not None:
            if _lst[0] == '' or _lst[0][-1] == ':' or _lst[0][-1] == '：':
                st[st.index(_lst)][1] = st[st.index(_lst)][2] = ''
            if _lst[3] == '' or _lst[3][-1] == ':' or _lst[3][-1] == '：':
                st[st.index(_lst)][4] = st[st.index(_lst)][5] = ''

    # 修正用公式计算的行
    _st_0 = ['zero']
    _st_1 = ['zero']
    for _lst in st:
        _st_0.append(_lst[0])
        _st_1.append(_lst[3])
    _i = 5
    for _lst in st[5:]:
        # print(_lst[0], _lst)
        _i = _i + 1
        st[st.index(_lst)][1] = '=VLOOKUP(A' + str(_i) + ',TB!A:E,5,0)'
        st[st.index(_lst)][4] = '=VLOOKUP(D' + str(_i) + ',TB!A:E,5,0)'
        if _lst[0] == '流动资产合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('货币资金')) + ':B' + str(
                _st_0.index('其他流动资产')) + ')-B' + str(_st_0.index('其中：应收利息')) + '-B' + str(_st_0.index('应收股利'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('货币资金')) + ':C' + str(
                _st_0.index('其他流动资产')) + ')-C' + str(_st_0.index('其中：应收利息')) + '-C' + str(_st_0.index('应收股利'))
        if _lst[0] == '非流动资产合计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_0.index('发放贷款和垫款')) + ':B' + str(_st_0.index('其他非流动资产'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_0.index('发放贷款和垫款')) + ':C' + str(_st_0.index('其他非流动资产'))
        if _lst[0] is None:
            st[st.index(_lst)][1] = ''
            st[st.index(_lst)][2] = ''
        if _lst[0] == '资产总计':
            st[st.index(_lst)][1] = '=B' + str(_st_0.index('流动资产合计')) + '+B' + str(_st_0.index('非流动资产合计'))
            st[st.index(_lst)][2] = '=C' + str(_st_0.index('流动资产合计')) + '+C' + str(_st_0.index('非流动资产合计'))
        if _lst[3] == '流动负债合计':
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_1.index('短期借款')) + ':E' + str(
                _st_1.index('其他流动负债')) + ')-E' + str(_st_1.index('其中：应付利息')) + '-E' + str(_st_1.index('应付股利'))
            st[st.index(_lst)][5] = '=SUM(F' + str(_st_1.index('短期借款')) + ':F' + str(
                _st_1.index('其他流动负债')) + ')-F' + str(_st_1.index('其中：应付利息')) + '-F' + str(_st_1.index('应付股利'))
        if _lst[3] == '非流动负债合计':
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_1.index('保险合同准备金')) + ':E' + str(
                _st_1.index('其他非流动负债')) + ')-E' + str(
                _st_1.index('其中：优先股(负债)')) + '-E' + str(_st_1.index('永续债(负债)'))
            st[st.index(_lst)][5] = '=SUM(F' + str(_st_1.index('保险合同准备金')) + ':F' + str(
                _st_1.index('其他非流动负债')) + ')-F' + str(
                _st_1.index('其中：优先股(负债)')) + '-F' + str(_st_1.index('永续债(负债)'))
        if _lst[3] == '负债合计':
            st[st.index(_lst)][4] = '=E' + str(_st_1.index('流动负债合计')) + '+E' + str(_st_1.index('非流动负债合计'))
            st[st.index(_lst)][5] = '=F' + str(_st_1.index('流动负债合计')) + '+F' + str(_st_1.index('非流动负债合计'))
        if _lst[3] == '归属于母公司所有者权益合计':
            st[st.index(_lst)][4] = '=SUM(E' + str(_st_1.index('实收资本(股本)')) + ':E' + str(
                _st_1.index('未分配利润')) + ',-E' + str(_st_1.index('其中：优先股(权益)')) + ',-E' + str(
                _st_1.index('永续债(权益)')) + ',-2*E' + str(_st_1.index('减：库存股'))
            st[st.index(_lst)][5] = '=SUM(F' + str(_st_1.index('实收资本(股本)')) + ':F' + str(
                _st_1.index('未分配利润')) + ',-F' + str(_st_1.index('其中：优先股(权益)')) + ',-F' + str(
                _st_1.index('永续债(权益)')) + ',-2*F' + str(_st_1.index('减：库存股'))
        if _lst[3] == '股东权益合计':
            st[st.index(_lst)][4] = '=E' + str(_st_1.index('归属于母公司所有者权益合计')) + '+E' + str(_st_1.index('少数股东权益'))
            st[st.index(_lst)][5] = '=F' + str(_st_1.index('归属于母公司所有者权益合计')) + '+F' + str(_st_1.index('少数股东权益'))
        if _lst[3] == '负债和股东权益总计':
            st[st.index(_lst)][4] = '=E' + str(_st_1.index('负债合计')) + '+E' + str(_st_1.index('股东权益合计'))
            st[st.index(_lst)][5] = '=F' + str(_st_1.index('负债合计')) + '+F' + str(_st_1.index('股东权益合计'))
    return st


def format_bs():
    wb.sheets('BS').range('D1:D1').api.ColumnWidth = 30
    wb.sheets('BS').range('B1:C1').api.ColumnWidth = 23
    wb.sheets('BS').range('E1:F1').api.ColumnWidth = 23
    wb.sheets('BS').range('A:F').api.Font.Name = "Arial"
    wb.sheets('BS').range('A1:F80').api.Style = "Comma"
    wb.sheets('BS').range('A4:F52').api.Borders(7).LineStyle = 1
    wb.sheets('BS').range('A4:F52').api.Borders(8).LineStyle = 1
    wb.sheets('BS').range('A4:F52').api.Borders(9).LineStyle = 1
    wb.sheets('BS').range('A4:F52').api.Borders(10).LineStyle = 1
    wb.sheets('BS').range('A4:F52').api.Borders(11).LineStyle = 1
    wb.sheets('BS').range('A4:F52').api.Borders(12).LineStyle = 1
    wb.sheets('BS').range('A1:F1').api.HorizontalAlignment = 7
    wb.sheets('BS').range('A21:F21').api.Font.Bold = True
    wb.sheets('BS').range('D35:F36').api.Font.Bold = True
    wb.sheets('BS').range('A51:F52').api.Font.Bold = True
    for lst in range(15):
        wb.sheets('BS').range('A' + str(lst + 6)).api.IndentLevel = 1
    for lst in [14, 15]:
        wb.sheets('BS').range('A' + str(lst)).api.IndentLevel = 2
    for lst in range(19):
        wb.sheets('BS').range('A' + str(lst + 23)).api.IndentLevel = 1
    for lst in range(15):
        wb.sheets('BS').range('D' + str(lst + 6)).api.IndentLevel = 1
    for lst in [16, 17]:
        wb.sheets('BS').range('D' + str(lst)).api.IndentLevel = 2
    for lst in range(12):
        wb.sheets('BS').range('D' + str(lst + 23)).api.IndentLevel = 1
    for lst in [26, 27]:
        wb.sheets('BS').range('D' + str(lst)).api.IndentLevel = 2
    for lst in range(11):
        wb.sheets('BS').range('D' + str(lst + 38)).api.IndentLevel = 1
    for lst in [40, 41]:
        wb.sheets('BS').range('D' + str(lst)).api.IndentLevel = 2
    return


# ----------------------------BS Over-------------------------------------------
# ----------------------------IS Start-------------------------------------------
def fill_is():
    st = wb.sheets('preIS').range('A1:C70').value
    # 修正无需计算的行
    for _lst in st[5:]:
        if _lst[0] is not None:
            if _lst[0] == '' or _lst[0][-1] == ':' or _lst[0][-1] == '：':
                st[st.index(_lst)][1] = st[st.index(_lst)][2] = ''

    # 修正用公式计算的行
    _st_1 = ['zero']
    for _lst in st:
        _st_1.append(_lst[0])
    _i = 4
    for _lst in st[4:]:
        # print(_lst[0], _lst)
        _i = _i + 1
        st[st.index(_lst)][1] = '=VLOOKUP(A' + str(_i) + ',TB!A:E,5,0)'
        if _lst[0] == '一、营业总收入':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('其中：营业收入')) + ':B' + str(_st_1.index('手续费及佣金收入'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('其中：营业收入')) + ':C' + str(_st_1.index('手续费及佣金收入'))
        if _lst[0] == '二、营业总成本':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('其中：营业成本')) + ':B' + str(_st_1.index('财务费用'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('其中：营业成本')) + ':C' + str(_st_1.index('财务费用'))
        if _lst[0] == '三、营业利润（亏损以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('一、营业总收入')) + '-B' + str(
                _st_1.index('二、营业总成本')) + '+SUM(B' + str(_st_1.index('加：其他收益')) + ':B' + str(
                _st_1.index('资产处置收益（损失以“-”号填列）')) + ')-B' + str(_st_1.index('其中：对联营企业和合营企业的投资收益')) + '-B' + str(
                _st_1.index('以摊余成本计量的金融资产终止确认收益'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('一、营业总收入')) + '-C' + str(
                _st_1.index('二、营业总成本')) + '+SUM(C' + str(_st_1.index('加：其他收益')) + ':C' + str(
                _st_1.index('资产处置收益（损失以“-”号填列）')) + ')-C' + str(_st_1.index('其中：对联营企业和合营企业的投资收益')) + '-C' + str(
                _st_1.index('以摊余成本计量的金融资产终止确认收益'))
        if _lst[0] == '四、利润总额（亏损总额以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('三、营业利润（亏损以“－”号填列）')) + '+B' + str(
                _st_1.index('加：营业外收入')) + '-B' + str(_st_1.index('减：营业外支出'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('三、营业利润（亏损以“－”号填列）')) + '+C' + str(
                _st_1.index('加：营业外收入')) + '-C' + str(_st_1.index('减：营业外支出'))
        if _lst[0] == '五、净利润（净亏损以“－”号填列）':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('四、利润总额（亏损总额以“－”号填列）')) + '-B' + str(
                _st_1.index('减：所得税费用'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('四、利润总额（亏损总额以“－”号填列）')) + '-C' + str(
                _st_1.index('减：所得税费用'))

    return st


def format_is():
    wb.sheets('IS').range('A1:A1').api.ColumnWidth = 55
    wb.sheets('IS').range('B1:C1').api.ColumnWidth = 20
    wb.sheets('IS').range('A:C').api.Font.Name = "Arial"
    wb.sheets('IS').range('A1:C80').api.Style = "Comma"
    wb.sheets('IS').range('A4:C70').api.Borders(7).LineStyle = 1
    wb.sheets('IS').range('A4:C70').api.Borders(8).LineStyle = 1
    wb.sheets('IS').range('A4:C70').api.Borders(9).LineStyle = 1
    wb.sheets('IS').range('A4:C70').api.Borders(10).LineStyle = 1
    wb.sheets('IS').range('A4:C70').api.Borders(11).LineStyle = 1
    wb.sheets('IS').range('A4:C70').api.Borders(12).LineStyle = 1
    wb.sheets('IS').range('A1:C1').api.HorizontalAlignment = 7
    wb.sheets('IS').range('A5:C5').api.Font.Bold = True
    wb.sheets('IS').range('A10:C10').api.Font.Bold = True
    wb.sheets('IS').range('A36:C36').api.Font.Bold = True
    wb.sheets('IS').range('A36:C36').api.Font.Bold = True
    wb.sheets('IS').range('A39:C39').api.Font.Bold = True
    wb.sheets('IS').range('A41:C41').api.Font.Bold = True
    wb.sheets('IS').range('A48:C48').api.Font.Bold = True
    wb.sheets('IS').range('A65:C65').api.Font.Bold = True
    wb.sheets('IS').range('A68:C68').api.Font.Bold = True
    for lst in range(63):
        wb.sheets('IS').range('a' + str(lst + 6)).api.IndentLevel = 1
    for lst in [10, 36, 39, 41, 48, 65, 68, 42, 45, 50, 56]:
        wb.sheets('IS').range('a' + str(lst)).api.IndentLevel = 0
    for lst in [28, 29]:
        wb.sheets('IS').range('a' + str(lst)).api.IndentLevel = 2
    return


# ----------------------------IS Over-------------------------------------------
# ----------------------------CF Start-------------------------------------------
def fill_cf():
    st = wb.sheets('preCF').range('A1:C70').value
    # 修正无需计算的行
    for _lst in st[5:]:
        if _lst[0] is not None:
            if _lst[0] == '' or _lst[0][-1] == ':' or _lst[0][-1] == '：':
                st[st.index(_lst)][1] = st[st.index(_lst)][2] = ''
    # 修正用公式计算的行
    _st_1 = ['zero']
    for _lst in st:
        _st_1.append(_lst[0])
    _i = 4
    for _lst in st[4:]:
        # print(_lst[0], _lst)
        if _lst[0] == '经营活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('销售商品、提供劳务收到的现金')) + ':B' + str(
                _st_1.index('收到其他与经营活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('销售商品、提供劳务收到的现金')) + ':C' + str(
                _st_1.index('收到其他与经营活动有关的现金'))
        if _lst[0] == '经营活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('购买商品、接受劳务支付的现金')) + ':B' + str(
                _st_1.index('支付其他与经营活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('购买商品、接受劳务支付的现金')) + ':C' + str(
                _st_1.index('支付其他与经营活动有关的现金'))
        if _lst[0] == '经营活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('经营活动现金流入小计')) + '-B' + str(
                _st_1.index('经营活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('经营活动现金流入小计')) + '-C' + str(
                _st_1.index('经营活动现金流出小计'))
        if _lst[0] == '投资活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('收回投资收到的现金')) + ':B' + str(
                _st_1.index('收到其他与投资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('收回投资收到的现金')) + ':C' + str(
                _st_1.index('收到其他与投资活动有关的现金'))
        if _lst[0] == '投资活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':B' + str(
                _st_1.index('支付其他与投资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('购建固定资产、无形资产和其他长期资产支付的现金')) + ':C' + str(
                _st_1.index('支付其他与投资活动有关的现金'))
        if _lst[0] == '投资活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('投资活动现金流入小计')) + '-B' + str(
                _st_1.index('投资活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('投资活动现金流入小计')) + '-C' + str(
                _st_1.index('投资活动现金流出小计'))
        if _lst[0] == '筹资活动现金流入小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('取得借款收到的现金')) + ':B' + str(
                _st_1.index('收到其他与筹资活动有关的现金')) + ',B' + str(_st_1.index('吸收投资收到的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('取得借款收到的现金')) + ':C' + str(
                _st_1.index('收到其他与筹资活动有关的现金')) + ',C' + str(_st_1.index('吸收投资收到的现金'))
        if _lst[0] == '筹资活动现金流出小计':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('偿还债务支付的现金')) + ':B' + str(
                _st_1.index('支付其他与筹资活动有关的现金'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('偿还债务支付的现金')) + ':C' + str(
                _st_1.index('支付其他与筹资活动有关的现金'))
        if _lst[0] == '筹资活动产生的现金流量净额':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('筹资活动现金流入小计')) + '-B' + str(
                _st_1.index('筹资活动现金流出小计'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('筹资活动现金流入小计')) + '-C' + str(
                _st_1.index('筹资活动现金流出小计'))
        if _lst[0] == '五、现金及现金等价物净增加额':
            st[st.index(_lst)][1] = '=SUM(B' + str(_st_1.index('经营活动产生的现金流量净额')) + ',B' + str(
                _st_1.index('投资活动产生的现金流量净额')) + ',B' + str(_st_1.index('筹资活动产生的现金流量净额')) + ',B' + str(
                _st_1.index('四、汇率变动对现金及现金等价物的影响'))
            st[st.index(_lst)][2] = '=SUM(C' + str(_st_1.index('经营活动产生的现金流量净额')) + ',C' + str(
                _st_1.index('投资活动产生的现金流量净额')) + ',C' + str(_st_1.index('筹资活动产生的现金流量净额')) + ',C' + str(
                _st_1.index('四、汇率变动对现金及现金等价物的影响'))
        if _lst[0] == '六、期末现金及现金等价物余额':
            st[st.index(_lst)][1] = '=B' + str(_st_1.index('五、现金及现金等价物净增加额')) + '+B' + str(
                _st_1.index('加：期初现金及现金等价物余额'))
            st[st.index(_lst)][2] = '=C' + str(_st_1.index('五、现金及现金等价物净增加额')) + '+C' + str(
                _st_1.index('加：期初现金及现金等价物余额'))

    return st


def format_cf():
    wb.sheets('CF').range('A1:C80').api.Font.Size = 11
    wb.sheets('CF').range('A1:A80').api.ColumnWidth = 55
    wb.sheets('CF').range('B1:C1').api.ColumnWidth = 25
    wb.sheets('CF').range('A:C').api.Font.Name = "Arial"
    wb.sheets('CF').range('A1:C80').api.Style = "Comma"
    wb.sheets('CF').range('A4:C60').api.Borders(7).LineStyle = 1
    wb.sheets('CF').range('A4:C60').api.Borders(8).LineStyle = 1
    wb.sheets('CF').range('A4:C60').api.Borders(9).LineStyle = 1
    wb.sheets('CF').range('A4:C60').api.Borders(10).LineStyle = 1
    wb.sheets('CF').range('A4:C60').api.Borders(11).LineStyle = 1
    wb.sheets('CF').range('A4:C60').api.Borders(12).LineStyle = 1
    wb.sheets('CF').range('A1:C1').api.HorizontalAlignment = 7
    wb.sheets('CF').range('A5:C5').api.Font.Bold = True
    wb.sheets('CF').range('A19:C19').api.Font.Bold = True
    wb.sheets('CF').range('A30:C32').api.Font.Bold = True
    wb.sheets('CF').range('A38:C38').api.Font.Bold = True
    wb.sheets('CF').range('A44:C46').api.Font.Bold = True
    wb.sheets('CF').range('A51:C51').api.Font.Bold = True
    wb.sheets('CF').range('A55:C58').api.Font.Bold = True
    wb.sheets('CF').range('A60:C60').api.Font.Bold = True
    for lst in range(55):
        wb.sheets('CF').range('a' + str(lst + 6)).api.IndentLevel = 1
    for lst in [19, 30, 31, 32, 38, 44, 45, 46, 51, 55, 56, 57, 58, 60]:
        wb.sheets('CF').range('a' + str(lst)).api.IndentLevel = 0
    return
