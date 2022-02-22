from macro import Macro


def run_macro():
    macro = Macro()
    ws = macro.open_ws()
    macro.clean_columns(ws)
    macro.check_columns_order(ws)  # * used after _clean_columns because initial order does not matter.
    macro.clean_rows_from_art_ressource(ws)
    macro.remove_not_anchor_products(ws)
    macro.loop_ankers_sum_lengths(ws)
    macro.save_excel()
    macro.calculate_consumption_loop()
    macro.put_number_to_new_sheet()


if __name__ == "__main__":
    run_macro()
