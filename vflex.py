import pandas as pd
from fuzzywuzzy import fuzz

def load_excel():
    file_one = pd.read_excel('test_excel1.xlsx', dtype=str).values.tolist()
    file_two = pd.read_excel('test_excel2.xlsx', dtype=str).values.tolist()
    return file_one, file_two

def set_search_col(file_one):
    print([x+1 for x in range(len(file_one[0]))])
    print(file_one[0])
    index = input('Which column to search for?')
    return [x[int(index) - 1] for x in file_one]

def set_reference_col(file_two):
    print([x+1 for x in range(len(file_two[0]))])
    print(file_two[0])
    return (int(input('Which column to search in?')) - 1)

def set_offset(file_two):
    print([x+1 for x in range(len(file_two[0]))])
    print(file_two[0])
    offset = input('Which offset range? (Ex. "2" or "2-5")').replace(' ','')
    while True:
        try:
            if '-' not in offset:
                offset = (int(offset), int(offset) + 1)
            else:
                temp_off = offset.split('-')
                offset = (int(temp_off[0]), int(temp_off[1]) + 1)
            break
        except:
            offset = input('Please input range again (Ex. "2" or "4-6")').replace(' ','')
    return offset
    
def get_results(search_terms, ref_col, offset, file_two):
    ref_terms = [x[ref_col] for x in file_two]
    for search_val in search_terms:
        for ref_ind, ref_val in enumerate(ref_terms):
            if fuzz.ratio(search_val, ref_val) > 92:
                print(offset)
                yield file_two[ref_ind][offset[0]:offset[1]]

def main():
    file_one, file_two = load_excel()
    search_terms = set_search_col(file_one)
    search_terms = ['TestTwo']
    reference_col = set_reference_col(file_two)
    offset_selection = set_offset(file_two)
    output = [x for x in get_results(search_terms, reference_col, offset_selection, file_two)]
    print(output)

if __name__ == '__main__':
    main()

