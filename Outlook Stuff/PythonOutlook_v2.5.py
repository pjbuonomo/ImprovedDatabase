def parse_line(line):
    # Original patterns
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at)?"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at)?"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid)\s+/\s+(\d+\.\d+)\s+(offer)"

    # New pattern for special phrases in price
    special_phrase_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(offered|bid)\s*(?:@|at|-)?\s*(\*\*.*?\*\*)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    if re.match(size_first_pattern, line) or re.match(name_first_pattern, line):
        for match in re.finditer(size_first_pattern, line) or re.finditer(name_first_pattern, line):
            size, name, cusip, price, action = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5]
            entries.append({"Name": name.strip(), "Size": size if size else "", "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})

    elif re.match(dual_action_pattern, line):
        for match in re.finditer(dual_action_pattern, line):
            size, name, cusip, bid_price, offer_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    elif re.match(special_phrase_pattern, line):
        for match in re.finditer(special_phrase_pattern, line):
            size, name, cusip, action, special_phrase = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[6]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": special_phrase, "Error": ""})

    return entries if entries else [default_dict]
