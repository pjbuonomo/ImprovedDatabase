def parse_line(line):
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*|(?<=at )\*\*.*?\*\*)\s+(bid|offered|offer)"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*|(?<=at )\*\*.*?\*\*)\s+(bid|offered|offer)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*)\s+(bid)\s+/\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*)\s+(offer)"
    bid_for_pattern = r"(\d+\.\d+)\s+bid\s+for\s+([\w\s-]+?)\s+\((\w+)\)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    if re.match(size_first_pattern, line):
        for match in re.finditer(size_first_pattern, line):
            size, name, cusip, price, action = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[6]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(name_first_pattern, line):
        for match in re.finditer(name_first_pattern, line):
            name, cusip, price, action = match.groups()[0], match.groups()[1], match.groups()[2], match.groups()[3]
            entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(dual_action_pattern, line):
        for match in re.finditer(dual_action_pattern, line):
            size, name, cusip, bid_price, offer_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    elif re.match(bid_for_pattern, line):
        for match in re.finditer(bid_for_pattern, line):
            price, name, cusip = match.groups()[0], match.groups()[1], match.groups()[2]
            entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": "bid", "Price": price, "Error": ""})

    return entries if entries else [default_dict]
