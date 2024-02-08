"""This is an experimental util.
"""
import openpyxl


def is_aligned_h(sheetdata, sheetmath, addr, horizontal):
    """Verifies that the horizontal alignment of the specified cell is specified.

    Returns True if the horizontal alignment of the cell is equal to the given one.
    The horizontal aligment values are listed in the following:
    https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.alignment.html#openpyxl.styles.alignment.Alignment.horizontal
    指定されたセルの水平方向の配置が与えられた指示と同じかどうかをチェックします。
    指示する値は、openpyxlに従います。
    https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.alignment.html#openpyxl.styles.alignment.Alignment.horizontal

    Parameters:
    ----------
    sheetdata : Worksheet instance of the openpyxl
        whose book are opened with data_only set to True
    sheetmath : Worksheet instance of the openpyxl
        whose book are opened with data_only set to False
    addr: str
        Excel style cell address.
        example: "A3"
    horizontal : str
        Alignment
        example : 'left', 'right', 'center', 'justify', and None.

    Returns:
    ----------
        bool
    """
    tmpboolH = False
    if horizontal is None:
        if sheetdata[addr].alignment.horizontal is None:
            tmpboolH = True
    else:
        if sheetdata[addr].alignment.horizontal == horizontal:
            tmpboolH = True
    return tmpboolH


def is_aligned_v(sheetdata, sheetmath, addr, vertical):
    """Verifies that the vertical alignment of the specified cell is specified.

    Returns True if the vertical alignment of the cell is equal to the given one.
    The horizontal aligment values are listed in the following:
    https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.alignment.html#openpyxl.styles.alignment.Alignment.vertical
    指定されたセルの垂直方向の配置が与えられた指示と同じかどうかをチェックします。
    指示する値は、openpyxlに従います。
    https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.alignment.html#openpyxl.styles.alignment.Alignment.vertical

    Parameters:
    ----------
    sheetdata : Worksheet instance of the openpyxl
        whose book are opened with data_only set to True
    sheetmath : Worksheet instance of the openpyxl
        whose book are opened with data_only set to False
    addr: str
        Excel style cell address.
        example: "A3"
    vertical : str
        Alignment
        example :  'top', 'bottom', 'center', 'justify', and None.

    Returns:
    ----------
        bool
    """
    tmpboolV = False
    if vertical is None:
        if sheetdata[addr].alignment.vertical is None:
            tmpboolV = True
    else:
        if sheetdata[addr].alignment.vertical == vertical:
            tmpboolV = True

    return tmpboolV


def is_solidfill(sheetdata, sheetmath, addr):
    """Verifies that the specified cell is filled with some color.

    Returns True if the cell is filled with some color.
    指定したセルが「塗りつぶし」かどうかをチェックします。

    Parameters:
    ----------
    sheetdata : Worksheet instance of the openpyxl
        whose book are opened with data_only set to True
    sheetmath : Worksheet instance of the openpyxl
        whose book are opened with data_only set to False
    addr: str
        Excel style cell address.
        example: "A3"

    Returns:
    ----------
        bool
    """
    if sheetdata[addr].fill.patternType == "solid":
        return True
    else:
        return False


def is_numberformat(sheetdata, sheetmath, addr, number_format):
    """Number format. セルの表示形式

    # 指定したセルの「表示形式」がｘｘである
    # General 標準
    # >>> sheetData["E60"].number_format
    # 'General'
    #
    # Number (built-in) 数値（組み込み）
    # >>> sheetData["E61"].number_format
    # '0_);[Red]\\(0\\)'
    #
    # Percentage (built-in) パーセンテージ（組み込み）
    # >>> sheetData["E62"].number_format
    # '0%'
    #
    # Text (built-in) 文字列（組み込み）
    # >>> sheetData["E63"].number_format
    # '@'
    #
    # Number (0 digits after decimal point) 数値（小数点以下０桁）
    # >>> sheetData["E64"].number_format
    # '0_ '
    #
    # Number (1 digits after decimal point) 数値（小数点以下１桁）
    # >>> sheetData["E65"].number_format
    # '0.0_ '
    #
    # Number (2 digits after decimal point) 数値（小数点以下２桁）
    # >>> sheetData["E66"].number_format
    # '0.00_ '
    #
    # Number (3 digits after decimal point) 数値（小数点以下３桁）
    # >>> sheetData["E67"].number_format
    # '0.000_ '
    """
    if sheetdata[addr].number_format == number_format:
        return True
    else:
        return False

# 指定したセルの「表示形式」の取得


def get_numberformat(sheetdata, sheetmath, addr):
    return sheetdata[addr].number_format


# 指定したセルの「フォント」がｘｘである
# 作成せず

# 指定したセルの「罫線」がｘｘである
# 作成せず

# 指定したセル範囲の「条件付き書式」がある
# 作成せず

# グラフ
# 作成せず
