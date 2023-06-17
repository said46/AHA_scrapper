import os

def prepare_tag_for_search(tag: str, asset: str) -> str:
    if tag[-3] == '_':
        tag = tag[:-3]
    if tag[0:2] == '%%':
        tag = tag[3:]
    if tag[-1].isalpha():
        last_part_beginning_index = -4
    else:
        last_part_beginning_index = -3
    middle_part = tag[3:last_part_beginning_index]
    if middle_part[1:] in ('ICA', 'IA', 'I', 'IC'):
        middle_part = middle_part[0] + 'T'
    if middle_part == 'XZGV':
        middle_part = 'XZV'
    if middle_part == 'XGV':
        middle_part = 'XV'
    if middle_part == 'HGV':
        middle_part = 'HV'
    if middle_part in ('PDI', 'PDIA', 'PDIC', 'PDICA'):
        middle_part = 'PDT'
    if middle_part == 'PDSH':
        middle_part = 'PD*'
    # temp:
    #if middle_part == 'XX':
    #    middle_part = 'X*'          
    # temp end
    # temp: 
    # if middle_part == 'XZGC':
    #    middle_part = 'XGC'        
    #if middle_part == 'XZGO':
    #    middle_part = 'XGO'          
    # temp end
    result = '*' + asset[0:3] + '-' + tag[:3] + '-' + middle_part + '-' + tag[last_part_beginning_index:]
    return result 

os.system("cls")

initial_tag = '004PDIC206_30'
print(f'{initial_tag=}')
key0 = '3000'
print(prepare_tag_for_search(initial_tag, key0))