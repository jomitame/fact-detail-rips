from openpyxl.styles import Font, Border, Side, Alignment, PatternFill


def give_style(p_stylo):
    """
    method for give style to a cell

    :param p_stylo: an int that references a style
    :return: a dict with parameters: border, font, align, fill
    """

    my_stylo = {}
    if (p_stylo == 'title'):
        ft = Font(size=10, bold=True, italic=True, color="000000")
        bd = Border(left=Side(border_style="double"), right=Side(border_style="double"),
                    top=Side(border_style="double"), bottom=Side(border_style="double"))
        al = Alignment(horizontal='center', vertical='center')
        rl = PatternFill(fill_type="solid", fgColor="dddddd")
    elif (p_stylo == 'total'):
        ft = Font(size=10, bold=False, italic=False, color="000000")
        bd = Border(left=Side(border_style="none"), right=Side(border_style="none"),
                    top=Side(border_style="none"),
                    bottom=Side(border_style="none"))
        al = Alignment(horizontal='right', vertical='bottom')
        rl = PatternFill(fill_type="solid", fgColor="dddddd")
    elif(p_stylo == 'normal'):
        ft = Font(size=8, bold=False, italic=False, color="000000")
        bd = Border(left=Side(border_style="none"), right=Side(border_style="none"),
                    top=Side(border_style="none"),
                    bottom=Side(border_style="none"))
        al = Alignment(horizontal='left', vertical='bottom')
        rl = PatternFill(fill_type="none")

    my_stylo.update({'fuente': ft})
    my_stylo.update({'borde': bd})
    my_stylo.update({'alineacion': al})
    my_stylo.update({'relleno': rl})

    return my_stylo