def parse_line(line):
    size_first_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid|offered|offer)\s*(?:@|at|-)?\s*(\d*\.\d+)?"
    name_first_pattern = r"([\w\s-]+?)\s+\((\w+)\)\s+(\d*\.\d+)?\s*(bid|offered|offer)\s*(?:@|at|-)?\s*(\d+\.\d+)"
    dual_action_pattern = r"(\d+(\.\d+)?(mm|m|k))\s+([\w\s-]+?)\s+\((\w+)\)\s+(\d+\.\d+)\s+(bid)\s+/\s+(\d+\.\d+)\s+(offer)"

    default_dict = {"Name": "", "Size": "", "CUSIP": "", "Actions": "", "Price": "", "Error": line}

    entries = []

    # Size-First Format
    if re.match(size_first_pattern, line):
        for match in re.finditer(size_first_pattern, line):
            size, name, cusip, price, action, alt_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[6], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": action, "Price": price if price else alt_price, "Error": ""})

    # Name-First Format
    elif re.match(name_first_pattern, line):
        for match in re.finditer(name_first_pattern, line):
            name, cusip, alt_price, action, price = match.groups()[0], match.groups()[1], match.groups()[2], match.groups()[3], match.groups()[4]
            entries.append({"Name": name.strip(), "Size": "", "CUSIP": cusip, "Actions": action, "Price": price if price else alt_price, "Error": ""})

    # Dual-Action Format
    elif re.match(dual_action_pattern, line):
        for match in re.finditer(dual_action_pattern, line):
            size, name, cusip, bid_price, offer_price = match.groups()[0], match.groups()[3], match.groups()[4], match.groups()[5], match.groups()[7]
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "bid", "Price": bid_price, "Error": ""})
            entries.append({"Name": name.strip(), "Size": size, "CUSIP": cusip, "Actions": "offer", "Price": offer_price, "Error": ""})

    return entries if entries else [default_dict]

Please show in all offerings. Many thanks for the focus.
Alamo 2023-1 A (011395AJ9) bid at 102.50
Blue Sky 2023-1 (XS2728630596) bid at 100.15
Bonanza 2022-1 A (09785EAJ0) bid at 90.00
Bonanza 2023-1 A (09785EAK7) bid at 99.90
Citrus 2023-1 B (177510AM6) bid at 102.40
Easton 2024-1 A (27777AAA9) bid at 100.25
First Coast 2021-1 (31971CAA1) bid at 96.15
First Coast 2023-1 (31969UAA5) bid at 101.10
Galileo 2023-1 B (36354TAP7) bid at 100.25
Galileo 2023-1 A (36354TAN2) bid at 100.25
Hypatia 2023-1 A (44914CAC0) bid at 104.35
Hexagon 2023-1 A (428270AA0) bid at 100.50
Lightning 2023-1 A (532242AA2) bid at 106.30
Matterhorn 2022-I B (577092AQ2) bid at 98.50
Merna 2022-2A (59013MAF9) bid at 98.65
Merna 2023-2 A (59013MAJ1) bid at 104.35
Mona Lisa 2023-1 B (608800AG3) bid at 107.75
Montoya 2022-2 (613752AB0) bid at 108.60
Montoya 2024-1 A (613752AC8) bid at 101.00
Ocelot 2023-1 A (675951AA5) bid at 100.30
Residential Re 2023-2 5 (76090WAC4) bid at 100.45
Tailwind 2022-1 B (87403TAE6) bid at 96.90
Tailwind 2022-1 C (87403TAE) bid at 98.15
Titania 2021-1 A (888329AA7) bid at 100.40
Titania 2021-2 A (888329AB5) bid at 97.50
Titania 2023-1 A (888329AC3) bid at 108.50
Ursa 2023-1 C (90323WAM2) bid at 100.45
Ursa 2023-3 D (90323WAQ3) bid at 100.35
2.25mm Gateway 2023-3 A (36779CAF3) offered @ 107.10
3mm Tailwind 2022-1 C (87403TAF3) 99.15 bid / 99.80 offer
4mm Tailwind 2022-1 B (87403TAE6) offered @ 99.10
3mm Tailwind 2022-1 C (87403TAF3) 99.20 bid / 99.50 offer

Please show in all offerings. Many thanks for the focus.




Alamo 2023-1 A (011395AJ9) bid at 102.50
Blue Sky 2023-1 (XS2728630596) bid at 100.15
Bonanza 2022-1 A (09785EAJ0) bid at 90.00
Bonanza 2023-1 A (09785EAK7) bid at 99.90
Citrus 2023-1 B (177510AM6) bid at 102.40
Easton 2024-1 A (27777AAA9) bid at 100.25
First Coast 2021-1 (31971CAA1) bid at 96.15
First Coast 2023-1 (31969UAA5) bid at 101.10
Galileo 2023-1 B (36354TAP7) bid at 100.25
Galileo 2023-1 A (36354TAN2) bid at 100.25
Hypatia 2023-1 A (44914CAC0) bid at 104.35
Hexagon 2023-1 A (428270AA0) bid at 100.50
Lightning 2023-1 A (532242AA2) bid at 106.30
Matterhorn 2022-I B (577092AQ2) bid at 98.50
Merna 2022-2A (59013MAF9) bid at 98.65
Merna 2023-2 A (59013MAJ1) bid at 104.35
Mona Lisa 2023-1 B (608800AG3) bid at 107.75
Montoya 2022-2 (613752AB0) bid at 108.60
Montoya 2024-1 A (613752AC8) bid at 101.00
Ocelot 2023-1 A (675951AA5) bid at 100.30
Residential Re 2023-2 5 (76090WAC4) bid at 100.45
Tailwind 2022-1 B (87403TAE6) bid at 99
Tailwind 2022-1 C (87403TAE) bid at 99.20
Titania 2021-1 A (888329AA7) bid at 100.40
Titania 2021-2 A (888329AB5) bid at 97.50
Titania 2023-1 A (888329AC3) bid at 108.50
Ursa 2023-1 C (90323WAM2) bid at 100.45
Ursa 2023-3 D (90323WAQ3) bid at 100.35
5mm Tailwind 2022-1 C (87403TAF3) 99.20 bid / 99.50 offer
500k Res Re 2020-I 13 (76124AAB4) offered @ 98.30
5mm Merna 2022-1 (59013MAE2) offered @ 100.50
500k Res Re 2020-I 13 (76124AAB4) offered @ **BH trades and we care to buy more**
500k Riverfront 2021-1 A (76870YAD4) 97.90 bid / 98.60 offer
Axed for more...
4mm Tailwind 2022-1 B (87403TAE6) offered @ **BH trades**
2.25mm Gateway 2023-3 A (36779CAF3) offered @ 107.00
Please show in all offerings. Many thanks for the focus. Many bids higher.


Alamo 2023-1 A (011395AJ9) bid at 102.50
Blue Sky 2023-1 (XS2728630596) bid at 100.60
Bonanza 2022-1 A (09785EAJ0) bid at 90.00
Bonanza 2023-1 A (09785EAK7) bid at 99.90
Citrus 2023-1 B (177510AM6) bid at 102.40
Easton 2024-1 A (27777AAA9) bid at 100.25
First Coast 2021-1 (31971CAA1) bid at 96.15
First Coast 2023-1 (31969UAA5) bid at 101.10
Galileo 2023-1 B (36354TAP7) bid at 100.25
Galileo 2023-1 A (36354TAN2) bid at 100.25
Hypatia 2023-1 A (44914CAC0) bid at 104.35
Hexagon 2023-1 A (428270AA0) bid at 100.90
Lightning 2023-1 A (532242AA2) bid at 106.30
Matterhorn 2022-I B (577092AQ2) bid at 98.50
Merna 2022-2A (59013MAF9) bid at 98.65
Merna 2023-2 A (59013MAJ1) bid at 104.35
Mona Lisa 2023-1 B (608800AG3) bid at 107.75
Montoya 2022-2 (613752AB0) bid at 108.60
Montoya 2024-1 A (613752AC8) bid at 101.00
Ocelot 2023-1 A (675951AA5) bid at 100.30
Residential Re 2023-2 5 (76090WAC4) bid at 100.45
Tailwind 2022-1 B (87403TAE6) bid at 99
Tailwind 2022-1 C (87403TAE) bid at 99.20
Titania 2021-1 A (888329AA7) bid at 100.40
Titania 2021-2 A (888329AB5) bid at 97.50
Titania 2023-1 A (888329AC3) bid at 108.50
Ursa 2023-1 C (90323WAM2) bid at 100.45
Ursa 2023-3 D (90323WAQ3) bid at 100.35
5mm Merna 2022-1 (59013MAE2) offered @ 100.50
500k Riverfront 2021-1 A (76870YAD4) 97.90 bid / 98.60 offer
2.25mm Gateway 2023-3 A (36779CAF3) 102 bid / 107.00 offer
3mm Mystic 2021-2 B (62865LAC1) offered @ 98.10
3mm Gateway 2022-1 A (36779CAA4) offered @ 102.00
1m Purple Re 2023-1 A (74639NAA1) offered at 103.10
74.50 bid for 3264 Re 2022-1 (88577CAB7)
64.50 bid for Herbie 2021-1 A (42703VAE3)
Please let us know if you care to offer... Thanks
Axed to buy more...
3mm Gateway 2022-1 A (36779CAA4) offered @ **BH trades**
3mm Mystic 2021-2 B (62865LAC1) - 97.35 bid / 98.10 offer
