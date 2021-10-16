import string


def remove_items_from_list(lst, *args):
    for arg in args:
        if isinstance(arg, list):
            lst = [x for x in lst if x not in arg]
        else:
            lst = [x for x in lst if x != arg]
    return lst


def check_against_truth_threshold(list_of_bools, threshold=0.5):
    return list_of_bools.count(True) / len(list_of_bools) >= threshold


def generate_delimiters(suffix='', *args):
    outline_letters = [f'{x}.{suffix}' for x in string.ascii_letters]
    delimiter_lst = [x for x in args]
    return outline_letters + delimiter_lst
