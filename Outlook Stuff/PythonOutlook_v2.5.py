def parse_line(line):
    # Pattern for lines with size first
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at)?"
    # Pattern for lines with name first
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at)?"
    # Pattern for dual-action lines
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid)\s+/\s+(\d+\.\d+)\s+(offer)"
    # Pattern for lines with special phrases in the price field
    special_phrase_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(offered|bid)\s*(?:@|at|-)?\s*(\*\*.*?\*\*)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    # Check each pattern
    if re.match(size_first_pattern, line) or re.match(name_first_pattern, line):
        pattern = size_first_pattern if re.match(size_first_pattern, line) else name_first_pattern
        for match in re.finditer(pattern, line):
            groups = match.groups()
            size = groups[0] if pattern == size_first_pattern else ""
            name = groups[3] if pattern == size_first_pattern else groups[0]
            cusip = groups[4] if pattern == size_first_pattern else groups[1]
            price = groups[5] if pattern == size_first_pattern else groups[2]
            action = groups[6] if pattern == size_first_pattern else groups[3]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(dual_action_pattern, line):
        for match in re.finditer(dual_action_pattern, line):
            groups = match.groups()
            size, name, cusip, bid_price, offer_price = groups[0], groups[3], groups[4], groups[5], groups[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    elif re.match(special_phrase_pattern, line):
        for match in re.finditer(special_phrase_pattern, line):
            groups = match.groups()
            size, name, cusip, action, special_phrase = groups[0], groups[3], groups[4], groups[5], groups[6]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": special_phrase, "Error": ""})

    return entries if entries else [default_dict]
