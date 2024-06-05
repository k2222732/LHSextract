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
