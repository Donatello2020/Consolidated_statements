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
        if _file_name[-3:] == "xls":
            _xlsx_name = _file_name
            break
    if _xlsx_name == '':
        print('没有找到Excel表格,请检查当前目录:' + _path_c)
        sys.exit(0)

    # 链接工作簿
    workbook = xw.Book(os.path.join(_path_c, _xlsx_name))

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
                    '应付债券', '其中：优先股', '永续债', '租赁负债', '长期应付款', '长期应付职工薪酬', '预计负债', '递延收益', '递延所得税负债', '其它非流动负债',
                    '非流动负债合计', '负债合计', '所有者权益：', '实收资本(股本)', '其它权益工具', '其中：优先股', '永续债', '资本公积', '减：库存股', '其它综合收益',
                    '专项储备', '盈余公积', '一般风险准备', '未分配利润', '归属于母公司所有者权益合计', '少数股东权益', '股东权益合计', '负债和股东权益总计']:
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
        wb.sheets('preiS').range('a' + str(lst + 6)).api.IndentLevel = 1
    for lst in [10, 36, 39, 41, 48, 65, 68, 42, 45, 50, 56]:
        wb.sheets('preiS').range('a' + str(lst)).api.IndentLevel = 0
    for lst in [28, 29]:
        wb.sheets('preiS').range('a' + str(lst)).api.IndentLevel = 2
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
    match_3(st, '筹资活动现金流出小计', '=sum(B51:B54)')
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
    match_3(st, '筹资活动现金流出小计', '=sum(C51:C54)', 0, 2)
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
    print(st)
    return st


def format_precf():
    wb.sheets('preCF').range('A1:C80').api.Font.Size = 11
    wb.sheets('preCF').range('A1:A80').api.ColumnWidth = 55
    wb.sheets('preCF').range('B1:C1').api.ColumnWidth = 20
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
                                                     '非流动负债：'], ['保险合同准备金'], ['长期借款'], ['应付债券'], ['其中：优先股'], ['永续债'], [
                                                     '租赁负债'], ['长期应付款'], ['长期应付职工薪酬'], ['预计负债'], ['递延收益'], [
                                                     '递延所得税负债'], ['其他非流动负债'], ['非流动负债合计'], ['负债合计'], ['所有者权益：'], [
                                                     '实收资本(股本)'], ['其他权益工具'], ['其中：优先股'], ['永续债'], ['资本公积'], [
                                                     '减：库存股'], ['其他综合收益'], ['专项储备'], ['盈余公积'], ['一般风险准备'], ['未分配利润'], [
                                                     '归属于母公司所有者权益合计'], ['少数股东权益'], ['股东权益合计'], ['负债和股东权益总计'], [
                                                     '一、营业总收入'], ['其中：营业收入'], ['利息收入_'], ['已赚保费'], ['手续费及佣金收入'], [
                                                     '二、营业总成本'], ['其中：营业成本'], ['利息支出'], ['手续费及佣金支出'], ['退保金'], [
                                                     '赔付支出净额'], ['提取保险责任准备金净额'], ['保单红利支出'], ['分保费用'], ['税金及附加'], [
                                                     '销售费用'], ['管理费用'], ['研发费用'], ['财务费用'], ['其中：利息费用'], ['利息收入'], [
                                                     '加：其他收益'], ['投资收益（损失以“－”号填列）'], ['其中：对联营企业和合营企业的投资收益'], [
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
    st = [['摘要', '借贷方向', '一级科目', '二级科目', '借方金额', '贷方金额', '备注']]
    return st
