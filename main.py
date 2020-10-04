# coding=utf-8
import func as fc

wb = fc.wblink()
fc.new_sheet('preBS', '现金流量表')
fc.new_sheet('preIS', 'preBS')
fc.new_sheet('preCF', 'preIS')
fc.new_sheet('.Validation', 'preCF')
fc.new_sheet('AJE', '.Validation')
wb.sheets('.Validation').api.Visible = False
fc.new_sheet('TB', 'AJE')
fc.new_sheet('BS', 'TB')
fc.new_sheet('IS', 'BS')
fc.new_sheet('CF', 'IS')

# preBS
wb.sheets('preBS').range('A1').value = fc.fill_prebs()
fc.format_prebs()

# preIS
wb.sheets('preIS').range('A1').value = fc.fill_preis()
fc.format_preis()

# preCF
wb.sheets('preCF').range('A1').value = fc.fill_precf()
fc.format_precf()

# AJE
fc.fill_validation()
wb.sheets('AJE').clear()
wb.sheets('AJE').range('A1').value = fc.fill_aje()
# wb.save()
# wb.close()
# xlApp = Dispatch('Excel.Application')
# xlApp.Quit()
# wb = fc.wblink()
# wb.sheets('AJE').range('C2:C500').api.Validation.Delete()
# fc.set_validation(wb.sheets('AJE').range('C2:C500'))
wb.sheets('AJE').range('A4').api.EntireRow.Delete()
fc.format_aje()

# TB
wb.sheets('TB').range('A1').value = fc.fill_tb()

wb.save()
