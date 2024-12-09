import math
import os
import logging
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("excel_processing.log", mode='w'),
        logging.StreamHandler()
    ]
)
ACCURACY = 0.001
LIMIT = 10000
MODEL_COL = 'B'
PRICE_EUR_COL = 'C'
PRICE_HRN_COL = 'D'
PERCENT_HRN_COL = 'E'
WEIGHT_PRICE_HRN_COL = 'F'
ADJUST_SUM_HRN_COL = 'G'
ADJUST_SUM_EUR_COL = 'H'
START_COLUMN = 2
END_COLUMN = 8
START_DATA_ROW = 6
DATE_CELL = f"{MODEL_COL}{START_COLUMN + 1}"
EURO_RATE_CELL = f"{PRICE_EUR_COL}{START_COLUMN + 1}"
ADJUST_SUM_EUR_CELL = f"{ADJUST_SUM_EUR_COL}{START_COLUMN + 1}"
ADJUST_SUM_HRN_CELL = f"{ADJUST_SUM_HRN_COL}{START_COLUMN + 1}"

PALE_BLUE_FILL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
PALE_GREEN_FILL = PatternFill(start_color="98FB98", end_color="98FB98", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def _round(value, number):
    value = round(value, 3)
    return float(Decimal(str(value)).quantize(Decimal('1.' + '1' * number), rounding=ROUND_HALF_UP))

def count_non_empty_models(ws):
    count = 0
    for row in range(START_DATA_ROW, ws.max_row + 1):
        if ws[f'{MODEL_COL}{row}'].value is not None:  # Check if the cell is not empty
            count += 1
        else:
            break  # Stop counting at the first empty cell
    return count


def process_excel_file(file_path):
    # Load the Excel file with data_only=True to get cell values instead of formulas
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    try:
        # date = ws[date_cell].value
        euro_rate = ws[EURO_RATE_CELL].value
        target_sum_hrn = ws[ADJUST_SUM_HRN_CELL].value
        ws[ADJUST_SUM_HRN_CELL].fill = PALE_GREEN_FILL
        non_empty_rows = count_non_empty_models(ws)
        target_sum_eur = target_sum_hrn / euro_rate
        sum_price_models_euro = sum(ws[f'{PRICE_EUR_COL}{row}'].value for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows))
        recalculated_diff_euro = target_sum_eur - sum_price_models_euro
    except TypeError:
        logging.error(f"Error reading values from the worksheet '{file_path}'. Check that the data is correct.")
        return

    ws[ADJUST_SUM_EUR_CELL].value = _round(target_sum_eur, 2)
    ws[ADJUST_SUM_EUR_CELL].fill = PALE_BLUE_FILL

    if sum(ws[f'{PERCENT_HRN_COL}{row}'].value for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows)) != 100:
        logging.error(f"Total weight percent in column '{PERCENT_HRN_COL}' of file {file_path} does not equal 100%.")
        return

    for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows):
        price_euro = ws[f'{PRICE_EUR_COL}{row}'].value
        price_model_hrn = price_euro * euro_rate
        weight_percent = ws[f'{PERCENT_HRN_COL}{row}'].value

        sum_by_percent = round(recalculated_diff_euro, 2) * euro_rate * weight_percent / 100
        ws[f'{PRICE_HRN_COL}{row}'] = price_model_hrn
        ws[f'{ADJUST_SUM_HRN_COL}{row}'] = ws[f'{PRICE_HRN_COL}{row}'].value + sum_by_percent
        ws[f'{ADJUST_SUM_EUR_COL}{row}'] = ws[f'{ADJUST_SUM_HRN_COL}{row}'].value / euro_rate
        ws[f'{WEIGHT_PRICE_HRN_COL}{row}'] = abs(sum_by_percent)

    adjust_hrn_values(ws, non_empty_rows, euro_rate)

    sum_hrn = sum(_round(ws[f'{ADJUST_SUM_HRN_COL}{row}'].value, 2) for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows))
    sum_eur = sum(_round(ws[f'{ADJUST_SUM_EUR_COL}{row}'].value, 2) for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows))

    if _round(sum_hrn, 2) != _round(target_sum_hrn, 2) or _round(sum_eur, 2) != _round(target_sum_eur, 2):
        logging.error(
            f"Could not reach required accuracy for file: {file_path}. "
            f"Expected summary sum HRN: {target_sum_hrn}, Calculated: {_round(sum_hrn, 2)}. "
            f"Expected summary sum EUR: {target_sum_eur}, Calculated: {_round(sum_eur, 2)}."
        )

    summary_row = START_DATA_ROW + non_empty_rows
    ws[f'{ADJUST_SUM_HRN_COL}{summary_row}'] = _round(sum_hrn, 2)
    ws[f'{ADJUST_SUM_EUR_COL}{summary_row}'] = _round(sum_eur, 2)

    ws[f'{ADJUST_SUM_HRN_COL}{summary_row}'].fill = PALE_GREEN_FILL
    ws[f'{ADJUST_SUM_EUR_COL}{summary_row}'].fill = PALE_BLUE_FILL

    for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows):
        ws[f'{PRICE_HRN_COL}{row}'].value = _round(ws[f'{PRICE_HRN_COL}{row}'].value, 2)
        ws[f'{WEIGHT_PRICE_HRN_COL}{row}'].value = _round(ws[f'{WEIGHT_PRICE_HRN_COL}{row}'].value, 2)
        ws[f'{ADJUST_SUM_HRN_COL}{row}'].value = _round(ws[f'{ADJUST_SUM_HRN_COL}{row}'].value, 2)
        # ws[f'{ADJUST_SUM_HRN_COL}{row}'].value = ws[f'{ADJUST_SUM_HRN_COL}{row}'].value
        ws[f'{ADJUST_SUM_EUR_COL}{row}'].value = _round(ws[f'{ADJUST_SUM_EUR_COL}{row}'].value, 2)

    set_borders(ws, summary_row)

    try:
        wb.save(file_path)
    except PermissionError:
        logging.error(f"Unable to save the workbook: {file_path}. Check to see if it's open.")
        return

    logging.info(f"Processing of file {file_path} completed successfully.")


def adjust_hrn_values(ws, non_empty_rows, euro_rate):

    def sum_of_fractions(fractions):
        fraction_str = f"{fractions:.6f}"[2:]
        return sum(int(digit) for digit in fraction_str)

    def sort_fractions_by_digit_sum(fractions, asc):
        return sorted(fractions, key=sum_of_fractions, reverse=not asc)

    def fill_values_holder_list(lst, asc=True):
        fractions = [math.modf(value)[0] for value in lst]
        sorted_fractions = sort_fractions_by_digit_sum(fractions, asc)

        holder = []
        for index, value in enumerate(lst):
            current_fraction = math.modf(value)[0]
            fraction_index = sorted_fractions.index(current_fraction)
            holder.append(ValuesHolder(fraction_index, index, value))

        return holder

    def calculate_values(col, cell, rounded=False):
        if rounded:
            values = [_round(ws[f'{col}{r}'].value, 2) for r in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows)]
        else:
            values = [ws[f'{col}{r}'].value for r in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows)]
        target_sum = ws[cell].value
        current_sum = sum(values)
        diff = target_sum - current_sum
        return values, target_sum, current_sum, diff

    values_hrn, target_sum_hrn, current_sum_hrn, diff_hrn = calculate_values(ADJUST_SUM_HRN_COL, ADJUST_SUM_HRN_CELL)
    values_eur, target_sum_eur, current_sum_eur, diff_eur = calculate_values(ADJUST_SUM_EUR_COL, ADJUST_SUM_EUR_CELL, True)

    values_helper = ValuesHelper(fill_values_holder_list(values_hrn, diff_hrn > 0))

    counter = 0
    while abs(diff_hrn) > ACCURACY or LIMIT > counter:
        current_value = values_helper.next()
        increment = ACCURACY if diff_hrn > 0 else -ACCURACY
        current_value_eur = ws[f'{ADJUST_SUM_EUR_COL}{START_DATA_ROW + values_helper.index()}'].value

        while LIMIT > counter:
            counter += 1

            new_value = current_value + increment
            old_value_eur = _round(current_value / euro_rate, 2)
            new_value_eur = _round(new_value / euro_rate, 2)
            updated_sum_hrn = current_sum_hrn - current_value + new_value
            updated_sum_eur = current_sum_eur - old_value_eur + new_value_eur

            if abs(old_value_eur > new_value_eur) == 0 and _round(abs(diff_eur), 2) == 0 and abs(updated_sum_hrn - target_sum_hrn) > abs(diff_hrn):
                break

            current_value = new_value
            current_value_eur = new_value_eur
            current_sum_hrn = updated_sum_hrn
            current_sum_eur = updated_sum_eur
            diff_hrn = target_sum_hrn - current_sum_hrn
            diff_eur = target_sum_eur - current_sum_eur

        values_helper.set_value_by_index(values_helper.index(), current_value)
        ws[f'{ADJUST_SUM_HRN_COL}{START_DATA_ROW + values_helper.index()}'].value = current_value
        ws[f'{ADJUST_SUM_EUR_COL}{START_DATA_ROW + values_helper.index()}'].value = current_value_eur

    for row in range(START_DATA_ROW, START_DATA_ROW + non_empty_rows):
        ws[f'{ADJUST_SUM_HRN_COL}{row}'].fill = YELLOW_FILL

def set_borders(ws, summary_row):
    thick_border_left = Border(left=Side(style="thick"))
    thick_border_right = Border(right=Side(style="thick"))

    for row in range(START_DATA_ROW - 3, summary_row + 1):
        for col in range(START_COLUMN, START_COLUMN + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thick_border_left
        for col in range(END_COLUMN, END_COLUMN + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thick_border_right

    thin_border = Border(bottom=Side(style="thin", color="000000"))
    thick_border = Border(bottom=Side(style="thick", color="000000"))

    for col in range(START_COLUMN, END_COLUMN + 1):
        ws.cell(row=START_DATA_ROW - 4, column=col).border = ws.cell(row=START_DATA_ROW - 4, column=col).border + thin_border
        ws.cell(row=START_DATA_ROW - 1, column=col).border = ws.cell(row=START_DATA_ROW - 1, column=col).border + thin_border
        ws.cell(row=START_DATA_ROW - 2, column=col).border = ws.cell(row=START_DATA_ROW - 2, column=col).border + thick_border
        ws.cell(row=summary_row - 1, column=col).border = ws.cell(row=summary_row - 1, column=col).border + thin_border
        ws.cell(row=summary_row, column=col).border = ws.cell(row=summary_row, column=col).border + thick_border

class ValuesHolder:
    def __init__(self, fraction_index, index, value):
        self.fraction_index = fraction_index
        self.index = index
        self.value = value

class ValuesHelper:
    current_fraction = -1
    def __init__(self, values):
        self.values = values

    def index(self):
        for item in self.values:
            if item.fraction_index == self.current_fraction:
                return item.index

    def get_item_by_fraction_index(self, idx):
        for item in self.values:
            if item.fraction_index == idx:
                return item

    def set_value_by_index(self, idx, value):
        for item in self.values:
            if item.index == idx:
                item.value = value

    def next(self):
        if self.current_fraction == len(self.values) - 1:
            self.current_fraction = 0
        else:
            self.current_fraction += 1
        item = self.get_item_by_fraction_index(self.current_fraction)
        return item.value

def main():
    folder_path = os.getcwd()
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    if len(excel_files) == 0:
        logging.error("No Excel files found in the current directory.")
        return

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        process_excel_file(file_path)


if __name__ == "__main__":
    main()
