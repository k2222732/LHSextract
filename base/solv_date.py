from datetime import datetime, timedelta
#输入年月日还有后推的天数，返回后推后日期的年月日
#calculate_date_add("2023-06-24",190)
def calculate_date_add(date_in, days:int):
    if isinstance(date_in, str):
        date_object = datetime.strptime(date_in, "%Y-%m-%d")
    else:
        date_object = datetime.strptime(date_in._date_repr, "%Y-%m-%d")
    
    new_date_object = date_object + timedelta(days=days)
    
    date_out = new_date_object.strftime("%Y-%m-%d")
    
    return date_out


def calculate_date_sub(date_in:str, days:int):
    if isinstance(date_in, str):
        date_object = datetime.strptime(date_in, "%Y-%m-%d")
    else:
        date_object = datetime.strptime(date_in._date_repr, "%Y-%m-%d")
    
    new_date_object = date_object - timedelta(days=days)
    
    date_out = new_date_object.strftime("%Y-%m-%d")
    
    return date_out