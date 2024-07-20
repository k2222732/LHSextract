from base.waitclick import *
#验证列表内容
def validate_list_content(wait, xpath, value):
    can= wait_return_subelement_absolute(wait, time_w = 0.5, xpath=xpath)
    char_amountof_member =can.get_attribute("textContent")
    if char_amountof_member != value:
        return False
    elif char_amountof_member == value:
        return True