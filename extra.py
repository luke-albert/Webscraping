symbol_and_name_list = td[1].text.split()
symbol = symbol_and_name_list[0]
name = symbol_and_name_list[1]
current_price = td[2].text
percent_change = td[8].text

current_price_without_dollarsign = current_price.replace('$', '')
percent_change_without_sign = percent_change.replace('%', '')
number_as_percent = float(percent_change_without_sign) / 100
yesterday_price = float(
    current_price_without_dollarsign) / (1 + number_as_percent)
round_yesterday_price = round(yesterday_price, 2)

write_sheet.cell(i, 1).value = symbol
write_sheet.cell(i, 2).value = name
write_sheet.cell(i, 3).value = current_price
write_sheet.cell(i, 4).value = percent_change
write_sheet.cell(i, 5).value = '$' + str(round_yesterday_price)
i += 1
