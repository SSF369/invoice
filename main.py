import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
from docx2pdf import convert


import datetime
from datetime import datetime

x = datetime.now()
xdate = x.strftime("%d-%m-%Y")

def digit_to_word(Total):
    # Define mappings for digit to word conversion in Indian number system
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
    teens = ["", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    # Convert the number to words
    if 1 <= Total < 10:
        return ones[Total]  
    elif 10 <= Total < 20:
        return teens[Total - 10]  
    elif 20 <= Total < 100:
        return tens[Total // 10] + (" " + ones[Total % 10] if Total % 10 != 0 else "")  
    elif 100 <= Total < 1000:
        return ones[Total // 100] + " Hundred" + (" and " + digit_to_word(Total % 100) if Total % 100 != 0 else "")  
    elif 1000 <= Total < 100000:
        return digit_to_word(Total // 1000) + " Thousand" + (" " + digit_to_word(Total % 1000) if Total % 1000 != 0 else "")  
    elif 100000 <= Total < 10000000:
        return digit_to_word(Total // 100000) + " Lakh" + (" " + digit_to_word(Total % 100000) if Total % 100000 != 0 else "")  
    elif 10000000 <= Total < 1000000000:
        return digit_to_word(Total // 10000000) + " Crore" + (" " + digit_to_word(Total % 10000000) if Total % 10000000 != 0 else "")   
    elif Total == 1000000000:
        return "One Billion"
    else:
        return "Total out of range"
'''

if __name__ == "__main__":
    # Replace '123456789' with your desired number
    number = int(input("Enter a number (up to 1,000,000,000): "))
    
    if 0 <= number <= 1000000000:
        result = digit_to_word(number)
        print(f"{number} in words: {result}")
    else:
        print("Number out of range (0-1,000,000,000)")
'''


def clear_item():
	#sno_spinbox.delete(0,tkinter.END)
	#sno_spinbox.insert(0,'1')
	#sno = 0
	desc_entry.delete(0,tkinter.END)
	hsncode_entry.delete(0,tkinter.END)
	bags_spinbox.delete(0,tkinter.END)
	bags_spinbox.insert(0,'1')
	weight_entry.delete(0,tkinter.END)
	rate_entry.delete(0,tkinter.END)

invoice_list = []

def add_item():
	sno = 1
	desc = desc_entry.get()
	hsncode = hsncode_entry.get()
	bags = int(bags_spinbox.get())
	weight = int(weight_entry.get())
	rate = int(rate_entry.get())
	line_amount = (rate*weight)/100
	#sno += 1
	invoice_item = [sno,desc,hsncode,bags,weight,rate,line_amount]
	tree.insert('',0,values=invoice_item)
	clear_item()
	invoice_list.append(invoice_item)
	

def new_invoice():
	broker_name_entry.delete(0,tkinter.END)
	invoice_no_entry.delete(0,tkinter.END)
	supply_date_entry.delete(0,tkinter.END)
	transport_entry.delete(0,tkinter.END)
	vehicle_number_entry.delete(0,tkinter.END)
	party_name_entry.delete(0,tkinter.END)
	party_address_entry.delete(0,tkinter.END)
	gstin_entry.delete(0,tkinter.END)
	state_entry.delete(0,tkinter.END)
	code_entry.delete(0,tkinter.END)
	clear_item()
	tree.delete(*tree.get_children())
	invoice_list.clear()

def generate_invoice():
	doc=DocxTemplate("Template.docx")
	#sno += 1
	broker_name = broker_name_entry.get()
	invoice_no = invoice_no_entry.get()
	invoice_date = xdate
	date_of_supply = supply_date_entry.get()
	transport_mode = transport_entry.get()
	vehicle_number = vehicle_number_entry.get()
	party_name = party_name_entry.get()
	party_address = party_address_entry.get()
	gstin = gstin_entry.get()
	state = state_entry.get()
	code = code_entry.get() 
	Total = sum(item[6] for item in invoice_list)
	total_in_words = digit_to_word((int(Total)))
	doc.render({ 'Broker_Name' : broker_name, 'Invoice_No' : invoice_no, 'Invoice_Date' : invoice_date, 'Date_Of_Supply' : date_of_supply, 'Transport_Mode' : transport_mode,'Vehicle_Number' : vehicle_number, 'Party_Name' : party_name, 'Party_Address' : party_address, 'Party_GSTIN' : gstin, 'Party_State' : state, 'Code' : code, 'Total' : Total, 'invoice_list' : invoice_list, 'Total_in_Words' : total_in_words + ' Rupees Only.'})
	doc_name = 'NT-'+ invoice_no + ".docx"

	doc.save(doc_name)
	
	new_invoice()

window = tkinter.Tk()
window.title("Invoice Generator")

frame = tkinter.Frame(window)
frame.pack(padx=20,pady=10)

broker_name_label = tkinter.Label(frame,text = "Broker Name")
broker_name_label.grid(row=0, column =0)
broker_name_entry = tkinter.Entry(frame)
broker_name_entry.grid(row=1, column =0)

invoice_no_label = tkinter.Label(frame,text = "Invoice Number")
invoice_no_label.grid(row=0, column =1)
invoice_no_entry = tkinter.Entry(frame)
invoice_no_entry.grid(row=1, column =1)

supply_date_label = tkinter.Label(frame,text = "Date Of Supply")
supply_date_label.grid(row=2, column =0)
supply_date_entry = tkinter.Entry(frame)
supply_date_entry.grid(row=3, column =0)

transport_label = tkinter.Label(frame,text = "Transport Mode")
transport_label.grid(row=2, column =1)
transport_entry = tkinter.Entry(frame)
transport_entry.grid(row=3, column =1)

vehicle_number_label = tkinter.Label(frame,text = "Vehicle Number")
vehicle_number_label.grid(row=4, column =0)
vehicle_number_entry = tkinter.Entry(frame)
vehicle_number_entry.grid(row=5, column =0)

party_address_label = tkinter.Label(frame,text = "Party Address")
party_address_label.grid(row=4, column =1)
party_address_entry = tkinter.Entry(frame)
party_address_entry.grid(row=5, column =1)

party_name_label = tkinter.Label(frame,text = "Party Name")
party_name_label.grid(row=6, column =0)
party_name_entry = tkinter.Entry(frame)
party_name_entry.grid(row=7, column =0)

gstin_label = tkinter.Label(frame,text = "GSTIN")
gstin_label.grid(row=6, column =1)
gstin_entry = tkinter.Entry(frame)
gstin_entry.grid(row=7, column =1)

state_label = tkinter.Label(frame,text = "State")
state_label.grid(row=8, column =0)
state_entry = tkinter.Entry(frame)
state_entry.grid(row=9, column =0)

code_label = tkinter.Label(frame,text = "Code")
code_label.grid(row=8, column =1)
code_entry = tkinter.Entry(frame)
code_entry.grid(row=9, column =1)

#sno_label = tkinter.Label(frame,text = "S.No")
#sno_label.grid(row=10, column =0)
#sno_spinbox = tkinter.Spinbox(frame, from_=0, to=1000000)
#sno_spinbox.grid(row=11, column =0)

desc_label = tkinter.Label(frame,text = "Product Desciption")
desc_label.grid(row=10, column =1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=11, column =1)

hsncode_label = tkinter.Label(frame,text = "Hsn Code")
hsncode_label.grid(row=12, column =0)
hsncode_entry = tkinter.Entry(frame)
hsncode_entry.grid(row=13, column =0)

bags_label = tkinter.Label(frame,text = "Bags")
bags_label.grid(row=12, column =1)
bags_spinbox = tkinter.Spinbox(frame, from_=0, to=1000000)
bags_spinbox.grid(row=13, column =1)


weight_label = tkinter.Label(frame,text = "Weight")
weight_label.grid(row=14, column =0)
weight_entry = tkinter.Entry(frame)
weight_entry.grid(row=15, column =0)

rate_label = tkinter.Label(frame,text = "Rate")
rate_label.grid(row=14, column =1)
rate_entry = tkinter.Entry(frame)
rate_entry.grid(row=15, column =1)

add_item_button = tkinter.Button(frame,text="Add Item",command = add_item)
add_item_button.grid(row=15, column=2, pady=5)

columns = ('sno', 'desc', 'hsncode', 'bags', 'weight', 'rate', 'amount')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('sno',text = 'S.No')
tree.heading('desc',text = 'Product Description')
tree.heading('hsncode',text='Hsn Code')
tree.heading('bags',text='Bags')
tree.heading('weight',text='Weight')
tree.heading('rate',text='Rate')
tree.heading('amount',text='Total Amuont')
tree.grid(row = 16, column = 0, columnspan = 4, padx = 20, pady = 10)

save_invoice_button = tkinter.Button(frame,text="Generate Invoice", command = generate_invoice)
save_invoice_button.grid(row = 17, column = 2, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame,text="New Invoice", command = new_invoice)
new_invoice_button.grid(row = 17, column = 1, sticky="news", padx=20, pady=5)


window.mainloop()
