def parse_line(line):
    # Regular expression patterns
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+([\d\.\*]+)\s+(bid|offered|offer)"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+([\d\.\*]+)\s+(bid|offered|offer)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+([\d\.\*]+)\s+(bid)\s+/\s+([\d\.\*]+)\s+(offer)"
    bid_for_pattern = r"(\d+\.\d+)\s+bid\s+for\s+([\w\s-]+?)\s+\((\w+)\)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    # Check each pattern
    for pattern in [size_first_pattern, name_first_pattern, dual_action_pattern, bid_for_pattern]:
        if re.match(pattern, line):
            for match in re.finditer(pattern, line):
                groups = match.groups()
                if pattern in [size_first_pattern, dual_action_pattern]:
                    size = groups[0]
                    name = groups[3]
                    cusip = groups[4]
                    price = groups[5]
                    action = groups[6] if pattern == size_first_pattern else "bid"
                    if pattern == dual_action_pattern:
                        offer_price = groups[7]
                        entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})
                else:
                    size = ""
                    name = groups[0] if pattern == name_first_pattern else groups[1]
                    cusip = groups[1] if pattern == name_first_pattern else groups[2]
                    price = groups[2] if pattern == name_first_pattern else groups[0]
                    action = groups[3] if pattern == name_first_pattern else "bid"

                entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price, "Error": ""})
            break  # Break after finding a match

    return entries if entries else [default_dict]
