from excel_utils import *


def parse_expression(expression):
    print expression
    operators = set('+-*/')
    quotes = set('"\'')
    quotes_open = False
    operations = []    # This holds the operators that are found in the string (left to right)
    operands = []   # this holds the non-operators that are found in the string (left to right)
    buff = []
    for character in expression:  # examine 1 character at a time
        if character in quotes:
            if quotes_open is True:
                quotes_open = False
            else:
                quotes_open = True
            buff.append(character)
        elif character in operators and quotes_open is False:
            # found an operator.  Everything we've accumulated in `buff` is
            # a single "number". Join it together and put it in `num_out`.
            operands.append(''.join(buff))
            buff = []
            operations.append(character)
        else:
            # not an operator.  Just accumulate this character in buff.
            buff.append(character)
    operands.append(''.join(buff))
    return operands, operations


def get_parsed_info_dict(expression):
    operands, operators = parse_expression(expression)
    print operands, operators
    nested_dict = {"operands": operands, "operands_size": len(operands),
                   "operators": operators, "operators_size": len(operators)}
    return nested_dict


def lookup_cell_definition(lookup_cell, sheet_to_lookup, **sheet_dict):
    # print lookup_cell
    lookup_cell_column, lookup_cell_row = cell_to_col_row(lookup_cell)
    initial_column = column_string_to_num(lookup_cell_column)
    # print initial_column
    initial_column_value = sheet_dict[sheet_to_lookup][row_col_to_cell(lookup_cell_row,initial_column)].value
    # print "working till here", initial_column_value
    # print initial_column_value[0] == '=' or initial_column_value is None, initial_column, initial_column_value
    # cell_definition = None
    print initial_column_value, '$' in initial_column_value
    if '$' in initial_column_value:
        cell_definition = lookup_cell_definition(initial_column_value[1:], sheet_to_lookup, **sheet_dict)
    else:
        while initial_column_value is None or initial_column_value[0] == '=':
            initial_column = initial_column - 1
            print lookup_cell_row, initial_column, initial_column_value
            initial_column_value = sheet_dict[sheet_to_lookup][row_col_to_cell(lookup_cell_row,initial_column)].value
            # print initial_column_value[0] == '=' or initial_column_value is None, initial_column, initial_column_value
        cell_definition = initial_column_value
    return cell_definition
