import ast
import builtins


# import openpyxl
# wb = openpyxl.load_workbook('py2excel.xlsx')
# sheet = wb['Sheet1']
# wb["Sheet1"]["a1"].value = '=LAMBDA(x, x+1)(3)'
# print(wb["Sheet1"]["a1"].value)
# wb.save('output.xlsx')


class LET:
    def __init__(self, name, value, expr):
        self.name = name
        self.value = value
        self.expr = expr

    def __repr__(self):
        return f'LET({self.name}, {self.value}, {self.expr})'

    def __str__(self):
        return f'LET("{self.name}", {self.value}, {self.expr})'


class LAMBDA:
    def __init__(self, param, expr):
        self.param = param
        self.expr = expr

    def __repr__(self):
        return f'LAMBDA({self.param}, {self.expr})'

    def __str__(self):
        return f'LAMBDA("{self.param}", {self.expr})'


# DO NOT EDIT
# hours_lost = 7
class JOIN:
    def __init__(self, sep, *args):
        self.sep = sep
        self.args = args

    def __repr__(self):
        return f'JOIN({self.sep}, {", ".join([repr(i) if isinstance(i, str) else str(i) for i in self.args])})'

    def __str__(self):
        return f'JOIN({self.sep}, {", ".join([repr(i) if isinstance(i, str) else str(i) for i in self.args])})'


class FUNC:
    def __init__(self, name, *expr):
        self.name = name
        self.expr = expr

    def __repr__(self):
        return f'{self.name}({", ".join([repr(i) if isinstance(i, str) else str(i) for i in self.expr])})'

    def __str__(self):
        return f'FUNC("{self.name}", {", ".join([repr(i) if isinstance(i, str) else str(i) for i in self.expr])})'


class STRING:
    def __init__(self, val):
        self.val = val

    def __repr__(self):
        return f"\"{self.val}\""

    def __str__(self):
        return f"STRING(\"{self.val}\")"


class IF:
    def __init__(self, cond, expr):
        self.cond = cond
        self.expr = expr

    def __repr__(self):
        return f'IF({self.cond}, {self.expr}, {STRING("")})'

    def __str__(self):
        return f'IF({self.cond}, {self.expr}, {STRING("")})'


def prettify(ast_tree_str, indent=4, file=None):
    ret = []
    stack = []
    in_string = False
    curr_indent = 0

    for i in range(len(ast_tree_str)):
        char = ast_tree_str[i]
        if in_string and char != '\'' and char != '"':
            ret.append(char)
        elif char == '(' or char == '[':
            ret.append(char)

            if i < len(ast_tree_str) - 1:
                next_char = ast_tree_str[i + 1]
                if next_char == ')' or next_char == ']':
                    curr_indent += indent
                    stack.append(char)
                    continue

            print(''.join(ret), file=file)
            ret.clear()

            curr_indent += indent
            ret.append(' ' * curr_indent)
            stack.append(char)
        elif char == ',':
            ret.append(char)

            print(''.join(ret), file=file)
            ret.clear()
            ret.append(' ' * curr_indent)
        elif char == ')' or char == ']':
            curr_indent -= indent
            ret.append('\n')
            ret.append(' ' * curr_indent)
            ret.append(char)
            stack.pop()
        elif char == '\'' or char == '"':

            if (len(ret) > 0 and ret[-1] == '\\') or (in_string and stack[-1] != char):
                ret.append(char)
                continue

            if len(stack) > 0 and stack[-1] == char:
                ret.append(char)
                in_string = False
                stack.pop()
                continue

            in_string = True
            ret.append(char)
            stack.append(char)
        elif char == ' ':
            pass
        else:
            ret.append(char)

    print(''.join(ret), file=file)


def set_con(expr, local, tree):
    if not local:
        global excel_code
        stack = [excel_code]
    else:
        stack = [tree]
    while stack:
        current = stack.pop()
        if hasattr(current, 'expr'):
            if current.expr == '__con':
                current.expr = expr
            else:
                stack.append(current.expr)


def general_bin_op(func, left_type, right_type, line, left, right, local, tree):
    match (left_type, right_type):
        case (builtins.int, builtins.int):
            set_con(
                LET(
                    line.targets[0].id,
                    FUNC(
                        func,
                        left,
                        right
                    ),
                    '__con'
                ), local, tree
            )
        case (builtins.int, builtins.tuple):
            set_con(
                LET(
                    line.targets[0].id,
                    FUNC(
                        func,
                        left,
                        right[0]
                    ),
                    '__con'
                ), local, tree
            )
        case (builtins.tuple, builtins.int):
            set_con(
                LET(
                    line.targets[0].id,
                    FUNC(
                        func,
                        left[0],
                        right
                    ),
                    '__con'
                ), local, tree
            )
        case (builtins.tuple, builtins.tuple):
            set_con(
                LET(
                    line.targets[0].id,
                    FUNC(
                        func,
                        left[0],
                        right[0]
                    ),
                    '__con'
                ), local, tree
            )


with open('code.pysx', 'r') as f:
    python_code = f.read()

python_code = ast.parse(python_code)
# prettify(ast.dump(python_code))


def default_tree():
    return LET('__z_combinator',
               LAMBDA(
                   '__f_y',
                   LET(
                       '__f_g',
                       LAMBDA(
                           '__f_x',
                           FUNC('__f_y',
                                LAMBDA(
                                    '__f_v',
                                    LET(
                                        '__f_xx',
                                        FUNC('__f_x',
                                             '__f_x'
                                             ),
                                        FUNC('__f_xx',
                                             '__f_v'
                                             )
                                    )
                                )
                                )
                       ),
                       FUNC('__f_g', '__f_g')
                   )
               ),
               '__con'
               )


excel_code = LET('__out', STRING(""), default_tree())
if_cons = []
while_count = 0


# So, so sorry for never nesters
def parse(code, local=False, tree=None):
    global while_count
    for line in code.body:
        match type(line):
            case ast.Expr:
                if isinstance(line.value, ast.Call):
                    if isinstance(line.value.func, ast.Name):
                        match line.value.func.id:
                            case 'print':
                                for arg in line.value.args:
                                    if isinstance(arg, ast.Constant):
                                        set_con(
                                            LET(
                                                '__out',
                                                JOIN(
                                                    STRING(" "),
                                                    '__out',
                                                    STRING(arg.value)
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    elif isinstance(arg, ast.Name):
                                        set_con(
                                            LET(
                                                '__out',
                                                JOIN(
                                                    STRING(" "),
                                                    '__out',
                                                    arg.id
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                if not line.value.keywords:
                                    set_con(
                                        LET(
                                            '__out',
                                            JOIN(
                                                STRING(" "),
                                                '__out',
                                                FUNC('CHAR', 10)
                                            ),
                                            '__con'
                                        ), local, tree
                                    )
                            case _:
                                set_con(
                                    FUNC(
                                        line.value.func.id,
                                        *line.value.args
                                    ), local, tree
                                )
            case ast.Assign:
                if len(line.targets) == 1:
                    if isinstance(line.value, ast.Constant):
                        set_con(
                            LET(
                                line.targets[0].id,
                                line.value.value if isinstance(line.value.value, int) else STRING(line.value.value),
                                '__con'
                            ), local, tree
                        )
                    elif isinstance(line.value, ast.BinOp):
                        match type(line.value.left):
                            case ast.Constant:
                                left_type = type(line.value.left.value)
                                left = line.value.left.value
                            case ast.Name:
                                left_type = tuple
                                left = line.value.left.id, 'var'
                            case _:
                                left_type = None
                                left = ...
                        match type(line.value.right):
                            case ast.Constant:
                                right_type = type(line.value.right.value)
                                right = line.value.right.value
                            case ast.Name:
                                right_type = tuple
                                right = line.value.right.id, 'var'
                            case _:
                                right_type = None
                                right = ...
                        match type(line.value.op):
                            case ast.Add:
                                match (left_type, right_type):
                                    case (builtins.str, builtins.str):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'CONCATENATE',
                                                    STRING(left),
                                                    STRING(right)
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.int, builtins.int):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'SUM',
                                                    left,
                                                    right
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.tuple, builtins.int):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'SUM',
                                                    left[0],
                                                    right
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.int, builtins.tuple):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'SUM',
                                                    left,
                                                    right[0]
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.tuple, builtins.tuple):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'SUM',
                                                    left[0],
                                                    right[0]
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.tuple, builtins.str):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'CONCATENATE',
                                                    left[0],
                                                    STRING(right)
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                                    case (builtins.str, builtins.tuple):
                                        set_con(
                                            LET(
                                                line.targets[0].id,
                                                FUNC(
                                                    'CONCATENATE',
                                                    STRING(left),
                                                    right[0]
                                                ),
                                                '__con'
                                            ), local, tree
                                        )
                            case ast.Sub:
                                general_bin_op('MINUS', left_type, right_type, line, left, right, local, tree)
                            case ast.Mult:
                                general_bin_op('PRODUCT', left_type, right_type, line, left, right, local, tree)
                            case ast.Div:
                                general_bin_op('DIVIDE', left_type, right_type, line, left, right, local, tree)
                            case ast.Mod:
                                general_bin_op('MOD', left_type, right_type, line, left, right, local, tree)
                            case ast.Pow:
                                general_bin_op('POW', left_type, right_type, line, left, right, local, tree)
                            case ast.LShift:
                                general_bin_op('BITLSHIFT', left_type, right_type, line, left, right, local, tree)
                            case ast.RShift:
                                general_bin_op('BITRSHIFT', left_type, right_type, line, left, right, local, tree)
                            case ast.BitOr:
                                general_bin_op('BITOR', left_type, right_type, line, left, right, local, tree)
                            case ast.BitAnd:
                                general_bin_op('BITAND', left_type, right_type, line, left, right, local, tree)
                            case ast.BitXor:
                                general_bin_op('BITXOR', left_type, right_type, line, left, right, local, tree)
            case ast.If:
                if True:
                    idx = len(if_cons)
                    if_cons.append(default_tree())
                    parse(line, True, if_cons[idx])
                    # set_con('_garbage', True, if_cons[idx])
                    # print(if_cons[idx])
                    comparison = ...
                    match type(line.test):
                        case ast.Constant:
                            match type(line.test.value):
                                case builtins.str:
                                    comparison = STRING(line.test.value)
                                case _:
                                    comparison = line.test.value
                        case ast.Name:
                            comparison = line.test.value
                        case ast.Compare:
                            def return_val(val):
                                match type(val):
                                    case ast.Constant:
                                        match type(val.value):
                                            case builtins.str:
                                                return STRING(val.value)
                                            case _:
                                                return val.value
                                    case ast.Name:
                                        return val.id
                            left, right = return_val(line.test.left), return_val(line.test.comparators[0])
                            compare_func = ...
                            match type(line.test.ops[0]):
                                case ast.Lt: compare_func = 'LT'
                                case ast.Eq: compare_func = 'EQ'
                                case ast.Gt: compare_func = 'GT'
                                case ast.LtE: compare_func = 'LTE'
                                case ast.GtE: compare_func = 'GTE'
                                case ast.NotEq: compare_func = 'NE'
                            comparison = FUNC(
                                compare_func,
                                left, right
                            )
                    set_con(IF(comparison, if_cons[idx]), local, tree)
            case ast.While:
                if True:
                    idx = len(if_cons)
                    if_cons.append(default_tree())
                    parse(line, True, if_cons[idx])
                    set_con('__out', True, if_cons[idx])
                    # print(if_cons[idx])
                    comparison = ...
                    match type(line.test):
                        case ast.Constant:
                            match type(line.test.value):
                                case builtins.str:
                                    comparison = STRING(line.test.value)
                                case _:
                                    comparison = line.test.value
                        case ast.Name:
                            comparison = line.test.value
                        case ast.Compare:
                            def return_val(val):
                                match type(val):
                                    case ast.Constant:
                                        match type(val.value):
                                            case builtins.str:
                                                return STRING(val.value)
                                            case _:
                                                return val.value
                                    case ast.Name:
                                        return val.id

                            left, right = return_val(line.test.left), return_val(line.test.comparators[0])
                            compare_func = ...
                            match type(line.test.ops[0]):
                                case ast.Lt:
                                    compare_func = 'LT'
                                case ast.Eq:
                                    compare_func = 'EQ'
                                case ast.Gt:
                                    compare_func = 'GT'
                                case ast.LtE:
                                    compare_func = 'LTE'
                                case ast.GtE:
                                    compare_func = 'GTE'
                                case ast.NotEq:
                                    compare_func = 'NE'
                            comparison = FUNC(
                                compare_func,
                                left, right
                            )
                    # set_con(IF(comparison, if_cons[idx]), local, tree)
                    set_con(
                        LET(
                            '_garbage',
                            FUNC(
                                '__z_combinator',
                                LAMBDA(
                                    '_garbage',
                                    LAMBDA(
                                        'cond',
                                        IF(
                                            'cond',
                                            LET(
                                                '_garbage_two',
                                                if_cons[idx],
                                                FUNC('_garbage', 'cond')
                                            )
                                        )
                                    )
                                )
                            ),
                            LET(
                                '_garbage_two',
                                FUNC('_garbage', comparison),
                                '__con'
                            )
                        ),
                        local, tree
                    )


parse(python_code)
set_con("'__out'", False, None)

LET.__str__ = LET.__repr__
LAMBDA.__str__ = LAMBDA.__repr__
JOIN.__str__ = JOIN.__repr__
STRING.__str__ = STRING.__repr__
FUNC.__str__ = FUNC.__repr__
IF.__str__ = IF.__repr__

excel_formula = '=' + str(excel_code).replace("'", "")
out_file = open('output.txt', 'w')
print(excel_formula, file=out_file)
out_file.close()
