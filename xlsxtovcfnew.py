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

class Contacts:

    #prefix,first_name,middle_name,last_name,suffix
    name_index         = [-1, -1, -1, -1, -1]
    num_index_labels   = []
    email_index_labels = []
    groups_index       = []
    #contact_file       = None
    def input_file(self):
        # input name of excel
        file = "50-sample-contacts.xlsx"
        xls = pd.ExcelFile(file)
        sheets = xls.sheet_names
        self.contacts_file = xls.parse(sheets[0])

    def print_cols(self):
        print(len(self.contacts_file),'is number of rows')
        for i,x in enumerate(self.contacts_file.columns):
            print(i,x)
        print("\n")
    def input_index(self,name_str):
        i = input("input for " + name_str+ " :")
        if i.isdigit():
            return int(i)
        elif i.isalnum():
            return i
        else:
            return -1

    def input_name(self):
        print("just press enter for nothing\n"
              "type index to extract from colounms\n"
              "type test to custom\n")
        for i,name in enumerate(["prefix","first_name", "middle_name", 'last_name', "suffix"]):
            self.name_index[i] = self.input_index(name)

    def input_nums(self):
        number_of_nums = int_input("number of phone numbers: ")
        for i in range(number_of_nums):
            num_index = int_input("index for phone number")
            num_label = input(f"label for {i + 1}.phone number")
            self.num_index_labels.append((num_index, num_label))

    def input_email(self):
        number_of_emails = int_input("number of email numbers: ")
        for i in range(number_of_emails):
            email_index = int_input("index for email number")
            email_label = input(f"label for {i+1}.email number")
            self.email_index_labels.append((email_index,email_label))
    def input_groups(self):
        number_of_groups= int_input("number of groups: ")
        if number_of_groups>0:
            print("just press enter for nothing\n"
                "type index to extract from colounms\n"
                "type test to custom\n")
            for i in range(number_of_groups):
                self.groups_index.append(input_index("{}.group".format(i)))

    def index_retriever(self, row_index, index):
        if type(index) == type(0):
            if index != -1:
                return str(self.contacts_file.iloc[row_index,index])
            else:
                return ""
        elif type(index) == type(""):
            return index
        else:
            exit(str(row_index, index))


    def build(self):
        vcf = ""
        for x in range(len(self.contacts_file)):
            vcf += "BEGIN:VCARD\nVERSION:3.0\n"


            # name
            full_name = ""
            for i in self.name_index:
                full_name += self.index_retriever(x, i)
            #debug
            print(full_name)
            vcf += "FN:" + full_name + '\n'

            # number+EMAIL
            temp = ''
            i = 0
            for index,label in self.num_index_labels:
                temp += f"items{i}.TEL:{self.index_retriever(x,index)}\n"
                temp += f"items{i}.X-ABLabel:{label}\n"
            for index,label in self.email_index_labels:
                temp += f"items{i}.TEL:{self.index_retriever(x,index)}\n"
                temp += f"items{i}.X-ABLabel:{label}\n"
            #debug
            print(temp)
            vcf += temp

            # categories
            temp = ""

            temp += "CATEGORIES:"
            for label in self.groups_index:
                temp += self.index_retriever(x,self.groups_index)
            temp += '\n'

            # end
            vcf += "END:VCARD\n\n"

        # saving file
        text_file = open("Export.vcf", "w", encoding="utf-8")  # Encoding utf-8 added
        text_file.write(vcf)
        text_file.close()
        print("Completed!")



contact = Contacts()
contact.input_file()
contact.print_cols()
contact.input_name()
contact.input_nums()
contact.input_email()
contact.input_groups()
contact.build()


text_file = open("Export.vcf", "w", encoding="utf-8")  # Encoding utf-8 added
text_file.write(vcf)
text_file.close()
print("Completed!")
