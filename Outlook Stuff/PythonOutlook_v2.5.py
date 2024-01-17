def parse_line(line):
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*|(?<=at )\*\*.*?\*\*)\s+(bid|offered|offer)"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*|(?<=at )\*\*.*?\*\*)\s+(bid|offered|offer)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*)\s+(bid)\s+/\s+(\d+\.\d+|(?<=@ )\*\*.*?\*\*)\s+(offer)"
    bid_for_pattern = r"(\d+\.\d+)\s+bid\s+for\s+([\w\s-]+?)\s+\((\w+)\)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    if re.match(size_first_pattern, line):
        matches = re.finditer(size_first_pattern, line)
        for match in matches:
            groups = match.groups()
            if len(groups) == 7:
                size, name, cusip, price, action = groups[0], groups[3], groups[4], groups[5], groups[6]
                entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(name_first_pattern, line):
        matches = re.finditer(name_first_pattern, line)
        for match in matches:
            groups = match.groups()
            if len(groups) == 4:
                name, cusip, price, action = groups[0], groups[1], groups[2], groups[3]
                entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(dual_action_pattern, line):
        matches = re.finditer(dual_action_pattern, line)
        for match in matches:
            groups = match.groups()
            if len(groups) == 8:
                size, name, cusip, bid_price, offer_price = groups[0], groups[3], groups[4], groups[5], groups[7]
                entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
                entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    elif re.match(bid_for_pattern, line):
        matches = re.finditer(bid_for_pattern, line)
        for match in matches:
            groups = match.groups()
            if len(groups) == 3:
                price, name, cusip = groups[0], groups[1], groups[2]
                entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": "bid", "Price": price, "Error": ""})

    return entries if entries else [default_dict]
