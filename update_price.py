#! python3
# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
from datetime import datetime
import argparse
import locale
import openpyxl
import os
import numpy as np

# %%
default_rules_path = 'update_rules.txt'


def get_rule(rule_str):
    if rule_str[0] == '+':
        def rule(price): return price+float(rule_str[1:])
    elif rule_str[0] == '-':
        def rule(price): return price-float(rule_str[1:])
    elif rule_str[0] == '*':
        def rule(price): return price*float(rule_str[1:])
    elif rule_str[0] == '/':
        def rule(price): return price/float(rule_str[1:])
    else:
        def rule(price): return price+float(rule_str)
    rule.__repr__ = rule_str
    return rule


def load_rules(rules_path=default_rules_path):
    rules = {'if': [], 'else': lambda price: price}
    with open(rules_path, 'r') as f:
        raw_rules = [line.replace(' ', '').replace('\t', '').split(',')
                     for line in f.read().split('\n')]
    status = []
    rule_funs = []
    for status_str, rule_str in raw_rules:
        if status_str == '-':
            rules['else'] = get_rule(rule_str)
        else:
            status.append(float(status_str))
            rule_funs.append(get_rule(rule_str))
    status_indice = np.array(status).argsort()
    rules['if'] = [(status[i], rule_funs[i]) for i in status_indice]
    return rules


rules = load_rules()


def update_price(value):
    for status, rule in rules['if']:
        if value <= status:
            return rule(value)
    return rules['else'](value)


# %%
def process(path, tar_path=None):
    file_name = os.path.basename(path)
    if not os.path.exists('output'):
        os.mkdir('output')

    wb = openpyxl.load_workbook(path)
    locale.setlocale(locale.LC_NUMERIC, 'C')

    for sheet in wb:
        if isinstance(sheet.iter_cols(), tuple):
            continue
        for col in sheet:
            for cell in col:
                find_price = False
                price = 0
                if isinstance(cell.value, int) or isinstance(cell.value, float):
                    find_price = True
                    price = cell.value
                # print(cell)
                if isinstance(cell.value, str):
                    try:
                        price = locale.atof(cell.value)
                        find_price = True
                        #print("price:", price)
                    except:
                        pass
                else:
                    # print(cell.value)
                    pass
                if find_price:
                    cell.value = update_price(price)
                #print(cell, cell.value)
    if tar_path is None:
        tar_path = os.path.join('output', '{}.{}.xlsx'.format('.'.join(file_name.split('.')[:-1]),
                                                              datetime.strftime(datetime.now(), '%m.%d')))
    wb.save(tar_path)


# %%
if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--out', type=str, default=None,
                        help='The save path')
    parser.add_argument('--config', type=str, default='update_rules.txt',
                        help='The rules path')
    parser.add_argument('file', type=str,
                        help='The input file which need to update')
    args = parser.parse_args()
    rules = load_rules(args.config)
    process(args.file, args.out)
