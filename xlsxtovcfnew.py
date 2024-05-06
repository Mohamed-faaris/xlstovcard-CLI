import pandas as pd
import os


def int_input(string):
    input_value = input(string)
    while not input_value.isdigit():
        print("input integer only")
        input_value = input(string)
    return int(input_value)


def int_e(string):
    input_value = input(string)
    try:
        return int(input_value)
    except:
        return -1

def input_index(name):
    i=input("input for" + name)
    if i.isdigit():
        return int(i)
    elif i.isalnum():
        return i
    else:
        return -1


class Contacts:

    #prefix,first_name,middle_name,last_name,suffix
    name_index         = [-1, -1, -1, -1, -1]
    num_index_labels   = []
    email_index_labels = []
    groups_index       = []

    def input_name(self):
        print("just press enter for nothing\n"
              "type index to extract from colounms\n"
              "type test to custom\n")
        for i,name in enumerate("prefix","first_name", "middle_name", 'last_name', "suffix"):
            self.name_index[i] = input_index(name)

    def input_nums(self):
        number_of_nums = int_input("number of phone numbers: ")
        for i in range(number_of_nums):
            num_index = int_input("index for phone number")
            num_label = input(f"label for {i + 1}.phone number")
            self.num_index_labels.append(tuple(num_index, num_label))

    def input_email(self):
        number_of_emails = int_input("number of email numbers: ")
        for i in range(number_of_emails):
            email_index = int_input("index for email number")
            email_label = input(f"label for {i+1}.email number")
            self.email_index_labels.append(tuple(email_index,email_label))
    def input_groups(self):
        number_of_groups= int_input("number of groups: ")
        if number_of_groups>0:
            print("just press enter for nothing\n"
                "type index to extract from colounms\n"
                "type test to custom\n")
            for i in range(number_of_groups):
                self.groups_index.append(input_index(f"{}.group".format(i)))


def build(contact):
    vcf = ""
    for x in range(len(contacts)):
        vcf += "BEGIN:VCARD\nVERSION:3.0\n"

        # name
        if option == 1:
            full_name = str(contacts[cols[name_index]][x])
            prefix = ""
            first_name = ""
            middle_name = ""
            last_name = ""
        else:  # elif option == 2:
            prefix = contacts[cols[prefix_index]][x] if prefix_index > 0 else ""
            first_name = contacts[cols[first_name_index]][x] if first_name_index > 0 else ""
            middle_name = contacts[cols[middle_name_index]][x] if middle_name_index > 0 else ""
            last_name = contacts[cols[last_name_index]][x] if last_name_index > 0 else ""
            full_name = prefix + " " + first_name + " " + middle_name + " " + last_name
        print(full_name)

        if option_suffix == 2:
            suffix = contacts[cols[suffix_index]][x]

        full_name += " " + suffix
        vcf += "FN:" + full_name + '\n'

        # number+email
        temp = ""
        for i in range(num_of_number + num_of_email):
            if i < num_of_email:
                temp += f"items{i}.EMAIL;TYPE=INTERNET:{contacts[cols[email_index[i]]][x]}\n"
                temp += f"items{i}.X-ABLabel:{email_labels[i]}\n"
            else:
                temp += f"items{i}.TEL:{contacts[cols[num_index[i - num_of_email]]][x]}\n"
                temp += f"items{i}.X-ABLabel:{num_labels[i - num_of_email]}\n"
        vcf += temp

        # categories
        if len(labels) != 0:
            labels_text = ""
            for label in labels:
                labels_text += label + ","
            vcf += "CATEGORIES:"
            vcf += labels_text[:-1]
            vcf += '\n'

        # end
        vcf += "END:VCARD\n\n"

    # saving file
    text_file = open("Export.vcf", "w", encoding="utf-8")  # Encoding utf-8 added
    text_file.write(vcf)
    text_file.close()
    print("Completed!")




# input name of excel
file = "50-sample-contacts.xlsx"
xls = pd.ExcelFile(file)
sheets = xls.sheet_names
if len(sheets) == 1:
    print(sheets[0], "is selected as contacts sheet since excel contains one sheet: ")
    contacts = xls.parse(sheets[0])
elif len(sheets) == 0:
    print("not sheets found")
    exit(1)
else:
    for i, sheet in enumerate(sheets):
        print(i + 1, sheet)
    i = int_input("input the index of sheet name: ") - 1
    contacts = xls.parse(sheets[i])
    print(sheets[i], "is selected")

print(len(contacts), "= rows in excel sheet")
cols = contacts.columns
for i, col in enumerate(cols):
    print(i + 1, col)

# name
option = int_input('1.if you'
                   'need name from single column(and a common suffix or suffix as any column)\n'
                   '2.if you need name as group of '
                   'prefix,first name,middle name,last name,suffix(any field can be empty)\n'
                   'choose: ')
if option == 1:
    name_index = int_input("enter the index for name: ") - 1
else:
    prefix_index = int_e("enter index for prefix: ") - 1
    first_name_index = int_e("enter index for first name: ") - 1
    middle_name_index = int_e("enter index for middle name: ") - 1
    last_name_index = int_e("enter index for last name: ") - 1

option_suffix: int = int_input("1.do need a common suffix\n"
                               "2.extract from a column\n"
                               "3.empty\n"
                               "choose: ")
if option_suffix == 1:
    suffix_index = -2
    suffix = input("common suffix: ")
elif option_suffix == 2:
    suffix_index = int_e("enter index for suffix: ") - 1
else:
    suffix_index = -1
    suffix = ""

# numbers
num_of_number = int_input("number of phone numbers: ")
num_index = []
num_labels = []
for x in range(num_of_number):
    num_labels.append(input("label for number: "))
    num_index.append(int_input("index for number: ") - 1)
    print(num_labels[x], ":", contacts.columns[num_index[x]])

# emails
num_of_email = int_input("number of email: ")
email_index = []
email_labels = []
for x in range(num_of_email):
    email_labels.append(input("label for email: "))
    email_index.append(int_input("index for email: ") - 1)
    print(email_labels[x], ":", contacts.columns[email_index[x]])

# labels
labels = []
num_of_groups = int_input("input the number of groups(labels): ")
for x in range(1, num_of_groups + 1):
    labels.append(input(f"{x}.group: "))

# DOB
# address
# job title
# company
# note


# build VCF
vcf = ""
for x in range(len(contacts)):
    vcf += "BEGIN:VCARD\nVERSION:3.0\n"

    # name
    if option == 1:
        full_name = str(contacts[cols[name_index]][x])
        prefix = ""
        first_name = ""
        middle_name = ""
        last_name = ""
    else:  # elif option == 2:
        prefix = contacts[cols[prefix_index]][x] if prefix_index > 0 else ""
        first_name = contacts[cols[first_name_index]][x] if first_name_index > 0 else ""
        middle_name = contacts[cols[middle_name_index]][x] if middle_name_index > 0 else ""
        last_name = contacts[cols[last_name_index]][x] if last_name_index > 0 else ""
        full_name = prefix + " " + first_name + " " + middle_name + " " + last_name
    print(full_name)

    if option_suffix == 2:
        suffix = contacts[cols[suffix_index]][x]

    full_name += " " + suffix
    vcf += "FN:" + full_name + '\n'

    # number+email
    temp = ""
    for i in range(num_of_number + num_of_email):
        if i < num_of_email:
            temp += f"items{i}.EMAIL;TYPE=INTERNET:{contacts[cols[email_index[i]]][x]}\n"
            temp += f"items{i}.X-ABLabel:{email_labels[i]}\n"
        else:
            temp += f"items{i}.TEL:{contacts[cols[num_index[i - num_of_email]]][x]}\n"
            temp += f"items{i}.X-ABLabel:{num_labels[i - num_of_email]}\n"
    vcf += temp

    # categories
    if len(labels) != 0:
        labels_text = ""
        for label in labels:
            labels_text += label + ","
        vcf += "CATEGORIES:"
        vcf += labels_text[:-1]
        vcf += '\n'

    # end
    vcf += "END:VCARD\n\n"

# saving file
text_file = open("Export.vcf", "w", encoding="utf-8")  # Encoding utf-8 added
text_file.write(vcf)
text_file.close()
print("Completed!")
