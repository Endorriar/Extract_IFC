# class data holder config

class data_holder():
    def __init__(self, category, prop_list):
        self.category = category
        self.prop_list = prop_list


class prop_holder():
    def __init__(self, pset, prop):
        self.pset = pset
        self.prop = prop


# helpfunction for filter

def filterConfig(_cat_name, _conf_lst):
    try:
        serach_res = [x for x in _conf_lst if x.category == _cat_name][0]
    except:
        serach_res = None
    return serach_res