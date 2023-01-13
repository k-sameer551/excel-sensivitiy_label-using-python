import win32com.client as win32


def excel_sensitivity_label(file_path):
    """open excel workbook"""
    xl = win32.Dispatch('Excel.Application')
    xl.Visible = False
    wb = xl.Workbooks.Open(file_path)
    set_sensitiviy_label(wb)
        
    
# setting sensitivity label
def set_sensitiviy_label(wbook):
    """set sensitivity label"""
    label = wbook.SensitivityLabel.CreateLabelInfo()
    label.AssignmentMethod = 1  #MsoAssignmentMethod.PRIVILEGED
    # label id and site id can be find in excel through vba sub routine
    label.LabelId = "a8a73c85-e524-44a6-bd58-7df7ef87be8f"
    label.SiteId = "6c15903a-880e-4e17-818a-6cb4f7935615"
    wbook.SensitivityLabel.SetLabel(label, label)


excel_sensitivity_label(file_path)