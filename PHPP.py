# Run some basic sanity tests for PHPP on LibreOffice.
# Look for any errors in any worksheet.

# Written by Ben Elliston, 2018

# 1. Create a scratch worksheet and call it 'Tests'.
# 2. Run this macro from the 'Tests' worksheet.
# 3. A table of worksheet names and error count will be generated.

def RunTests(*args):
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    controller = model.getCurrentController()
    active = controller.ActiveSheet

    sheets = [s for s in model.Sheets if s.getName() != 'Tests']
    for i, sheet in enumerate(sheets):
        cell = active.getCellByPosition(0, i)
        cell.setString(sheet.getName())
        cell = active.getCellByPosition(1, i)
        cell.setFormula("=SUM(IF(ISERROR('%s'.A1:ZZ1000)))" % sheet.getName())
    return None
