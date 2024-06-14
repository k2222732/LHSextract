#读写配置文件
import configparser



#拿取配置文件中的参数
def pick_config_param(configfilepath, str_class, str_param):
    '''
    configfilepath:配置文件的路径
    str_class:参数归属的大类的名称
    str_param:参数的名称
    '''
    config_file = configfilepath
    config = configparser.ConfigParser()
    config.read(config_file)
    result = config.get(str_class, str_param, fallback='')
    return result



#写入配置文件中的参数
def pick_config_param(configfilepath, str_class_name, str_param_name, str_param):
    '''
    configfilepath:配置文件的路径
    str_class:参数归属的大类的名称
    str_param:参数的名称
    '''
    # 创建配置解析器对象
    config = configparser.ConfigParser()
    # 读取现有的配置文件
    config.read(configfilepath)
    # 如果配置文件中不存在指定的大类，添加它
    if str_class_name not in config.sections():
        config.add_section(str_class_name)
    # 设置指定大类中的参数
    config.set(str_class_name, str_param_name, str_param)
    # 将修改后的配置写回文件
    with open(configfilepath, 'w') as configfile:
        config.write(configfile)



if __name__ == "__main__":
    pick_config_param()