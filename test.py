from tabulate import tabulate


# Function to parse and process the data
def process_data(raw_data):
    # Split the data by lines
    rows = raw_data.strip().split("\n")

    # Split the header (first line)
    headers = rows[0].split("\t")

    # Prepare the data as a list of dictionaries
    data = {}
    for row in rows[1:]:
        row_data = row.split("\t")
        category = row_data[0]  # The category name
        data[category] = dict(
            zip(headers[1:], row_data[1:]))  # Skip the category column and create the dictionary for each category

    return data


# Function to print the data in the desired format
def print_data(data):
    for category, details in data.items():
        # Format the output in the specified style
        print(f'"{category}": {details},')


raw_data = """Category	Total 2025	Percent to Total 2025	Apr 2024	May 2024	Jun 2024	Jul 2024	Aug 2024	Sep 2024
 Rooms ΑΛΒΑΝΙΑ	5	0.02%	0	0	3	0	0	2
 Rooms ΑΥΣΤΡΑΛΙΑ	9	0.04%	0	5	0	0	3	1
 Rooms ΑΥΣΤΡΙΑ	169	0.67%	4	28	102	8	23	4
 Rooms ΒΕΛΓΙΟ	80	0.32%	0	1	15	32	10	22
 Rooms ΒΟΡΕΙΑ ΜΑΚΕΔΟΝΙΑ	1836	7.26%	40	460	402	188	208	538
 Rooms ΒΟΣΝΙΑ-ΕΡΖΕΓΟΒΙΝΗ	3	0.01%	0	0	0	0	0	3
 Rooms ΒΟΥΛΓΑΡΙΑ	7194	28.44%	12	1026	1990	1100	638	2428
 Rooms ΓΑΛΛΙΑ	117	0.46%	0	5	28	36	36	12
 Rooms ΓΕΡΜΑΝΙΑ	958	3.79%	0	123	245	132	304	154
 Rooms ΓΕΩΡΓΙΑ	9	0.04%	0	0	0	0	4	5
 Rooms ΔΑΝΙΑ	28	0.11%	0	0	28	0	0	0
 Rooms ΕΛΒΕΤΙΑ	148	0.59%	0	25	66	29	24	4
 Rooms ΕΛΛΗΝΙΚΗ	9053	35.79%	2	503	1238	3101	3238	971
 Rooms Η.Π.Α.	113	0.45%	2	28	19	39	12	13
 Rooms ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ	170	0.67%	0	2	38	37	77	16
 Rooms ΙΡΛΑΝΔΙΑ	10	0.04%	0	0	3	0	0	7
 Rooms ΙΣΠΑΝΙΑ	3	0.01%	0	0	2	0	0	1
 Rooms ΙΣΡΑΗΛ	30	0.12%	0	7	6	9	7	1
 Rooms ΙΤΑΛΙΑ	1332	5.27%	1	18	116	391	735	71
 Rooms ΚΑΝΑΔΑΣ	33	0.13%	0	4	0	5	24	0
 Rooms ΚΟΣΟΒΟ	11	0.04%	0	2	9	0	0	0
 Rooms ΚΡΟΑΤΙΑ	6	0.02%	0	0	0	0	0	6
 Rooms ΚΥΠΡΟΣ	65	0.26%	0	0	4	18	38	5
 Rooms ΛΕΤΤΟΝΙΑ	8	0.03%	0	0	2	0	4	2
 Rooms ΛΙΘΟΥΑΝΙΑ	4	0.02%	0	4	0	0	0	0
 Rooms ΜΑΛΤΑ	4	0.02%	0	0	0	0	0	4
 Rooms ΜΟΛΔΑΒΙΑ	50	0.20%	0	12	0	8	6	24
 Rooms ΝΟΡΒΗΓΙΑ	8	0.03%	0	8	0	0	0	0
 Rooms ΟΛΛΑΝΔΙΑ	159	0.63%	0	15	42	34	39	29
 Rooms ΟΥΓΓΑΡΙΑ	117	0.46%	0	0	23	44	35	15
 Rooms ΟΥΚΡΑΝΙΑ	84	0.33%	2	11	14	13	3	41
 Rooms ΠΟΛΩΝΙΑ	124	0.49%	0	9	33	34	41	7
 Rooms ΠΟΡΤΟΓΑΛΙΑ	26	0.10%	0	4	0	9	9	4
 Rooms ΡΟΥΜΑΝΙΑ	820	3.24%	14	125	108	114	193	266
 Rooms ΡΩΣΙΑ	26	0.10%	0	6	0	11	0	9
 Rooms ΣΕΡΒΙΑ	1520	6.01%	41	264	295	347	128	445
 Rooms ΣΛΟΒΑΚΙΑ	71	0.28%	0	2	15	29	16	9
 Rooms ΣΛΟΒΕΝΙΑ	108	0.43%	0	3	8	78	11	8
 Rooms ΣΟΥΗΔΙΑ	32	0.13%	0	2	6	24	0	0
 Rooms ΤΟΥΡΚΙΑ	714	2.82%	2	23	271	173	116	129
 Rooms ΤΣΕΧΙΑ, ΔΗΜΟΚΡΑΤΙΑ ΤΗΣ	41	0.16%	0	0	14	27	0	0
Total Rooms Final 2024	25298	100.00%	120	2725	5145	6070	5982	5256
"""
data2 = {" Rooms ΑΛΒΑΝΙΑ": {'Total Final 2024': '5', 'Percent to Total Final 2024': '0.02%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                            'Jun Final 2024': '3', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '2'},
         " Rooms ΑΥΣΤΡΑΛΙΑ": {'Total Final 2024': '9', 'Percent to Total Final 2024': '0.04%', 'Apr Final 2024': '0', 'May Final 2024': '5',
                              'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '3', 'Sep 2024': '1'},
         " Rooms ΑΥΣΤΡΙΑ": {'Total Final 2024': '169', 'Percent to Total Final 2024': '0.67%', 'Apr Final 2024': '4', 'May Final 2024': '28',
                            'Jun Final 2024': '102', 'Jul Final 2024': '8', 'Aug Final 2024': '23', 'Sep 2024': '4'},
         " Rooms ΒΕΛΓΙΟ": {'Total Final 2024': '80', 'Percent to Total Final 2024': '0.32%', 'Apr Final 2024': '0', 'May Final 2024': '1',
                           'Jun Final 2024': '15', 'Jul Final 2024': '32', 'Aug Final 2024': '10', 'Sep 2024': '22'},
         " Rooms ΒΟΡΕΙΑ ΜΑΚΕΔΟΝΙΑ": {'Total Final 2024': '1836', 'Percent to Total Final 2024': '7.26%', 'Apr Final 2024': '40',
                                     'May Final 2024': '460', 'Jun Final 2024': '402', 'Jul Final 2024': '188', 'Aug Final 2024': '208',
                                     'Sep 2024': '538'},
         " Rooms ΒΟΣΝΙΑ-ΕΡΖΕΓΟΒΙΝΗ": {'Total Final 2024': '3', 'Percent to Total Final 2024': '0.01%', 'Apr Final 2024': '0',
                                      'May Final 2024': '0', 'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '0',
                                      'Sep 2024': '3'},
         " Rooms ΒΟΥΛΓΑΡΙΑ": {'Total Final 2024': '7194', 'Percent to Total Final 2024': '28.44%', 'Apr Final 2024': '12',
                              'May Final 2024': '1026', 'Jun Final 2024': '1990', 'Jul Final 2024': '1100', 'Aug Final 2024': '638',
                              'Sep 2024': '2428'},
         " Rooms ΓΑΛΛΙΑ": {'Total Final 2024': '117', 'Percent to Total Final 2024': '0.46%', 'Apr Final 2024': '0', 'May Final 2024': '5',
                           'Jun Final 2024': '28', 'Jul Final 2024': '36', 'Aug Final 2024': '36', 'Sep 2024': '12'},
         " Rooms ΓΕΡΜΑΝΙΑ": {'Total Final 2024': '958', 'Percent to Total Final 2024': '3.79%', 'Apr Final 2024': '0', 'May Final 2024': '123',
                             'Jun Final 2024': '245', 'Jul Final 2024': '132', 'Aug Final 2024': '304', 'Sep 2024': '154'},
         " Rooms ΓΕΩΡΓΙΑ": {'Total Final 2024': '9', 'Percent to Total Final 2024': '0.04%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                            'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '4', 'Sep 2024': '5'},
         " Rooms ΔΑΝΙΑ": {'Total Final 2024': '28', 'Percent to Total Final 2024': '0.11%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                          'Jun Final 2024': '28', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '0'},
         " Rooms ΕΛΒΕΤΙΑ": {'Total Final 2024': '148', 'Percent to Total Final 2024': '0.59%', 'Apr Final 2024': '0', 'May Final 2024': '25',
                            'Jun Final 2024': '66', 'Jul Final 2024': '29', 'Aug Final 2024': '24', 'Sep 2024': '4'},
         " Rooms ΕΛΛΗΝΙΚΗ": {'Total Final 2024': '9053', 'Percent to Total Final 2024': '35.79%', 'Apr Final 2024': '2',
                             'May Final 2024': '503', 'Jun Final 2024': '1238', 'Jul Final 2024': '3101', 'Aug Final 2024': '3238',
                             'Sep 2024': '971'},
         " Rooms Η.Π.Α.": {'Total Final 2024': '113', 'Percent to Total Final 2024': '0.45%', 'Apr Final 2024': '2', 'May Final 2024': '28',
                           'Jun Final 2024': '19', 'Jul Final 2024': '39', 'Aug Final 2024': '12', 'Sep 2024': '13'},
         " Rooms ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ": {'Total Final 2024': '170', 'Percent to Total Final 2024': '0.67%', 'Apr Final 2024': '0',
                                     'May Final 2024': '2', 'Jun Final 2024': '38', 'Jul Final 2024': '37', 'Aug Final 2024': '77',
                                     'Sep 2024': '16'},
         " Rooms ΙΡΛΑΝΔΙΑ": {'Total Final 2024': '10', 'Percent to Total Final 2024': '0.04%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                             'Jun Final 2024': '3', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '7'},
         " Rooms ΙΣΠΑΝΙΑ": {'Total Final 2024': '3', 'Percent to Total Final 2024': '0.01%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                            'Jun Final 2024': '2', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '1'},
         " Rooms ΙΣΡΑΗΛ": {'Total Final 2024': '30', 'Percent to Total Final 2024': '0.12%', 'Apr Final 2024': '0', 'May Final 2024': '7',
                           'Jun Final 2024': '6', 'Jul Final 2024': '9', 'Aug Final 2024': '7', 'Sep 2024': '1'},
         " Rooms ΙΤΑΛΙΑ": {'Total Final 2024': '1332', 'Percent to Total Final 2024': '5.27%', 'Apr Final 2024': '1', 'May Final 2024': '18',
                           'Jun Final 2024': '116', 'Jul Final 2024': '391', 'Aug Final 2024': '735', 'Sep 2024': '71'},
         " Rooms ΚΑΝΑΔΑΣ": {'Total Final 2024': '33', 'Percent to Total Final 2024': '0.13%', 'Apr Final 2024': '0', 'May Final 2024': '4',
                            'Jun Final 2024': '0', 'Jul Final 2024': '5', 'Aug Final 2024': '24', 'Sep 2024': '0'},
         " Rooms ΚΟΣΟΒΟ": {'Total Final 2024': '11', 'Percent to Total Final 2024': '0.04%', 'Apr Final 2024': '0', 'May Final 2024': '2',
                           'Jun Final 2024': '9', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '0'},
         " Rooms ΚΡΟΑΤΙΑ": {'Total Final 2024': '6', 'Percent to Total Final 2024': '0.02%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                            'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '6'},
         " Rooms ΚΥΠΡΟΣ": {'Total Final 2024': '65', 'Percent to Total Final 2024': '0.26%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                           'Jun Final 2024': '4', 'Jul Final 2024': '18', 'Aug Final 2024': '38', 'Sep 2024': '5'},
         " Rooms ΛΕΤΤΟΝΙΑ": {'Total Final 2024': '8', 'Percent to Total Final 2024': '0.03%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                             'Jun Final 2024': '2', 'Jul Final 2024': '0', 'Aug Final 2024': '4', 'Sep 2024': '2'},
         " Rooms ΛΙΘΟΥΑΝΙΑ": {'Total Final 2024': '4', 'Percent to Total Final 2024': '0.02%', 'Apr Final 2024': '0', 'May Final 2024': '4',
                              'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '0'},
         " Rooms ΜΑΛΤΑ": {'Total Final 2024': '4', 'Percent to Total Final 2024': '0.02%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                          'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '4'},
         " Rooms ΜΟΛΔΑΒΙΑ": {'Total Final 2024': '50', 'Percent to Total Final 2024': '0.20%', 'Apr Final 2024': '0', 'May Final 2024': '12',
                             'Jun Final 2024': '0', 'Jul Final 2024': '8', 'Aug Final 2024': '6', 'Sep 2024': '24'},
         " Rooms ΝΟΡΒΗΓΙΑ": {'Total Final 2024': '8', 'Percent to Total Final 2024': '0.03%', 'Apr Final 2024': '0', 'May Final 2024': '8',
                             'Jun Final 2024': '0', 'Jul Final 2024': '0', 'Aug Final 2024': '0', 'Sep 2024': '0'},
         " Rooms ΟΛΛΑΝΔΙΑ": {'Total Final 2024': '159', 'Percent to Total Final 2024': '0.63%', 'Apr Final 2024': '0', 'May Final 2024': '15',
                             'Jun Final 2024': '42', 'Jul Final 2024': '34', 'Aug Final 2024': '39', 'Sep 2024': '29'},
         " Rooms ΟΥΓΓΑΡΙΑ": {'Total Final 2024': '117', 'Percent to Total Final 2024': '0.46%', 'Apr Final 2024': '0', 'May Final 2024': '0',
                             'Jun Final 2024': '23', 'Jul Final 2024': '44', 'Aug Final 2024': '35', 'Sep 2024': '15'},
         " Rooms ΟΥΚΡΑΝΙΑ": {'Total Final 2024': '84', 'Percent to Total Final 2024': '0.33%', 'Apr Final 2024': '2', 'May Final 2024': '11',
                             'Jun Final 2024': '14', 'Jul Final 2024': '13', 'Aug Final 2024': '3', 'Sep 2024': '41'},
         " Rooms ΠΟΛΩΝΙΑ": {'Total Final 2024': '124', 'Percent to Total Final 2024': '0.49%', 'Apr Final 2024': '0', 'May Final 2024': '9',
                            'Jun Final 2024': '33', 'Jul Final 2024': '34', 'Aug Final 2024': '41', 'Sep 2024': '7'},
         " Rooms ΠΟΡΤΟΓΑΛΙΑ": {'Total Final 2024': '26', 'Percent to Total Final 2024': '0.10%', 'Apr Final 2024': '0', 'May Final 2024': '4',
                               'Jun Final 2024': '0', 'Jul Final 2024': '9', 'Aug Final 2024': '9', 'Sep 2024': '4'},
         " Rooms ΡΟΥΜΑΝΙΑ": {'Total Final 2024': '820', 'Percent to Total Final 2024': '3.24%', 'Apr Final 2024': '14', 'May Final 2024': '125',
                             'Jun Final 2024': '108', 'Jul Final 2024': '114', 'Aug Final 2024': '193', 'Sep 2024': '266'},
         " Rooms ΡΩΣΙΑ": {'Total Final 2024': '26', 'Percent to Total Final 2024': '0.10%', 'Apr Final 2024': '0', 'May Final 2024': '6',
                          'Jun Final 2024': '0', 'Jul Final 2024': '11', 'Aug Final 2024': '0', 'Sep 2024': '9'},
         " Rooms ΣΕΡΒΙΑ": {'Total Final 2024': '1520', 'Percent to Total Final 2024': '6.01%', 'Apr Final 2024': '41', 'May Final 2024': '264',
                           'Jun Final 2024': '295', 'Jul Final 2024': '347', 'Aug Final 2024': '128', 'Sep 2024': '445'},
         " Rooms ΣΛΟΒΑΚΙΑ": {'Total Final 2024': '71', 'Percent to Total Final 2024': '0.28%', 'Apr Final 2024': '0', 'May Final 2024': '2',
                             'Jun Final 2024': '15', 'Jul Final 2024': '29', 'Aug Final 2024': '16', 'Sep 2024': '9'},
         " Rooms ΣΛΟΒΕΝΙΑ": {'Total Final 2024': '108', 'Percent to Total Final 2024': '0.43%', 'Apr Final 2024': '0', 'May Final 2024': '3',
                             'Jun Final 2024': '8', 'Jul Final 2024': '78', 'Aug Final 2024': '11', 'Sep 2024': '8'},
         " Rooms ΣΟΥΗΔΙΑ": {'Total Final 2024': '32', 'Percent to Total Final 2024': '0.13%', 'Apr Final 2024': '0', 'May Final 2024': '2',
                            'Jun Final 2024': '6', 'Jul Final 2024': '24', 'Aug Final 2024': '0', 'Sep 2024': '0'},
         " Rooms ΤΟΥΡΚΙΑ": {'Total Final 2024': '714', 'Percent to Total Final 2024': '2.82%', 'Apr Final 2024': '2', 'May Final 2024': '23',
                            'Jun Final 2024': '271', 'Jul Final 2024': '173', 'Aug Final 2024': '116', 'Sep 2024': '129'},
         " Rooms ΤΣΕΧΙΑ, ΔΗΜΟΚΡΑΤΙΑ ΤΗΣ": {'Total Final 2024': '41', 'Percent to Total Final 2024': '0.16%', 'Apr Final 2024': '0',
                                           'May Final 2024': '0', 'Jun Final 2024': '14', 'Jul Final 2024': '27', 'Aug Final 2024': '0',
                                           'Sep 2024': '0'},
         "Total Rooms Final 2024": {'Total Final 2024': '25298', 'Percent to Total Final 2024': '100.00%', 'Apr Final 2024': '120',
                                    'May Final 2024': '2725', 'Jun Final 2024': '5145', 'Jul Final 2024': '6070', 'Aug Final 2024': '5982',
                                    'Sep 2024': '5256'}}
# Process the raw data
processed_data = process_data(raw_data)

# Print the processed table
print_data(processed_data)
