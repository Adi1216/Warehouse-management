#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
excel_path = 'warehouse.xlsx'
df = pd.read_excel(excel_path)

# function to remove item from the excel file

def remove_item(df,product_name):
    if product_name.upper() in df['Product'].values:
        df = df[df['Product'] != product_name.upper()]
        df = df.reset_index(drop=True)
        df.to_excel(excel_path,index = False)
        df = pd.read_excel(excel_path)
        return df
    else:
        print(f" Product '{product_name}' not found, Removal failed.")
        return df
#function to add new item to the excel file.
def add_item(df,serial,product_name,quantity,location,AirForceSerial,catalog):
    df = pd.read_excel(excel_path)
    if (serial == ''):
        serial = '0'
    if (AirForceSerial == ''):
        AirForceSerial = '0'
    if (catalog == ''):
        catalog = '0'
    if (product_name.upper() in df['Product'].values):
        return (f" Product '{product_name}' is already exist")
    else:
        new_row = {'Product' : product_name.upper(),
                   'Quantity': int(quantity),
                   'Location': location.upper(),
                   'SerialNumber' : serial.upper(),
                   'AirForceSerial' : AirForceSerial.upper(),
                   'catalog' : catalog.upper()}
        df = df.append(new_row,ignore_index =True)
        df.to_excel(excel_path,index=False)
        return (f" Product '{product_name}', has been inserted")
    
    
#function to update Quantity in the excel file.
def update_quantity(df,product_name,new_quantity):
    df = pd.read_excel(excel_path)
    if product_name.upper() in df['Product'].values:
        negativeCheck = df[df['Product'] == product_name.upper()]
        if (negativeCheck.Quantity.values[0] + int(new_quantity) < 0):
            return(f" Quantity Can't be under Zero \n Current Quantity of '{product_name}' is '{negativeCheck.Quantity.values[0]}'")
        else:
            df.loc[df['Product'] == product_name.upper(), 'Quantity'] += int(new_quantity)
            df.to_excel(excel_path,index=False)
    else:
        return (f"Product '{product_name}', not found, Updated failed.")
    product_row = df[df['Product'] == product_name.upper()]
    quantity = product_row['Quantity'].values[0]
    if (quantity == 0):
        remove_item(df,product_name)
        return (f" Product '{product_name}' found, and has been deleted beacuse the Quantity is 0")
    else:
        return (f"Product '{product_name}', has been Updated successfully with Quantity of '{quantity}'")
    
# function to serach every matching substring (return all)
def search_product(product,df):
    df = pd.read_excel(excel_path)
    matching_product = df[df['Product'].str.contains(product.upper(), case = False, na = False)]
    if (matching_product.size != 0):
        return (matching_product)
    else:
        return (matching_product)
excel_path = 'warehouse.xlsx'
df = pd.read_excel(excel_path)

class App():
    def __init__(self):
        self.root = tk.Tk()
#Change the Entire box Size\
        self.root.geometry("1250x370")
#name of the entire box
        self.root.title("Warehouse Managment")
#creating listbox  
        width = '100'
        height = '1'
        rowspan = '1'
        self.product_list = tk.Listbox(self.root, width = width, height = height)
        self.product_list.grid(row=2,column = 12, columnspan = 10, padx=10, pady= 10, rowspan = rowspan)
        self.search_label = tk.Label(self.root, text="Results", font = ("Helvetica", 20, 'bold'), foreground = 'brown')
        self.search_label.grid(row=1, column=12, columnspan=10, padx=10, pady=10)
        
#data insert title
        self.title_label = ttk.Label(self.root, text = "Enter Data", font = ("Helvetica", 16, 'bold'), foreground = 'blue')
        self.title_label.grid(row=0,column=0, columnspan=2,padx=10,pady= 10)
#update form title
        self.update_label = ttk.Label(self.root,text="Update Data", font=("Helvetica", 16, "bold"), foreground = 'blue')
        self.update_label.grid(row=0,column=5,columnspan=2,padx=10,pady=10)    
#search form title
        self.search_label = ttk.Label(self.root,text="Search Data", font=("Helvetica", 16, "bold"), foreground = 'blue')
        self.search_label.grid(row=0,column=10,columnspan=2,padx=10,pady=10)  
#product
        self.product_label = ttk.Label(self.root, text = 'Product', foreground="red")
        self.product_label.grid(row = 1 , column = 1, padx = 10 , pady = 10)
        self.product_entry = ttk.Entry(self.root)
        self.product_entry.grid(row=1,column=0,padx=10,pady=10)
#quantity
        self.quantity_label = ttk.Label(self.root, text = 'Quantity', foreground="red")
        self.quantity_label.grid(row = 2 , column = 1, padx = 10 , pady = 10)
        self.quantity_entry = ttk.Entry(self.root)
        self.quantity_entry.grid(row=2,column=0,padx=10,pady=10)
#location
        self.location_label = ttk.Label(self.root, text = 'Location', foreground="red")
        self.location_label.grid(row = 3 , column = 1, padx = 10 , pady = 10)
        self.location_entry = ttk.Entry(self.root)
        self.location_entry.grid(row=3,column=0,padx=10,pady=10)
#serial
        self.serial_label = ttk.Label(self.root, text = 'SerialNumber')
        self.serial_label.grid(row = 4 , column = 1, padx = 10 , pady = 10)
        self.serial_entry = ttk.Entry(self.root)
        self.serial_entry.grid(row=4,column=0,padx=10,pady=10)
        
#Airforce
        self.airforce_label = ttk.Label(self.root, text = 'ProductNumber')
        self.airforce_label.grid(row = 5 , column = 1, padx = 10 , pady = 10)
        self.airforce_entry = ttk.Entry(self.root)
        self.airforce_entry.grid(row=5,column=0,padx=10,pady=10)
#catalog
        self.catalog_label = ttk.Label(self.root, text = 'Catalog')
        self.catalog_label.grid(row = 6 , column = 1, padx = 10 , pady = 10)
        self.catalog_entry = ttk.Entry(self.root)
        self.catalog_entry.grid(row=6,column=0,padx=10,pady=10)
#product_to_update
        self.update_label = ttk.Label(self.root, text = 'Product Update')
        self.update_label.grid(row = 1 , column = 6, padx = 10 , pady = 10)
        self.update_entry = ttk.Entry(self.root)
        self.update_entry.grid(row=1,column=5,padx=10,pady=10)
#update amount [Add]
        self.new_quantity_label = ttk.Label(self.root, text = 'Quantity to add')
        self.new_quantity_label.grid(row = 2 , column = 6, padx = 10 , pady = 10)
        self.new_quantity_entry = ttk.Entry(self.root)
        self.new_quantity_entry.grid(row=2,column=5,padx=10,pady=10)
#update amount [sub]
        self.remove_quantity_label = ttk.Label(self.root, text = 'Quantity to remove')
        self.remove_quantity_label.grid(row = 3 , column = 6, padx = 10 , pady = 10)
        self.remove_quantity_entry = ttk.Entry(self.root)
        self.remove_quantity_entry.grid(row=3,column=5,padx=10,pady=10)
#product_to_search
        self.productToShow_label = ttk.Label(self.root, text = 'Search Product')
        self.productToShow_label.grid(row=1 ,column = 10, padx = 10 , pady = 10)
        self.productToShow_entry = ttk.Entry(self.root)
        self.productToShow_entry.grid(row=2,column=10,padx=10,pady=10)
#submit button
        self.data=[]
        self.submit_button = ttk.Button(self.root, text = "Confirm", command = self.submit_data)
        self.submit_button.grid(row = 7, column = 0 , columnspan =2, padx = 10 , pady = 10)
#update button
        self.update_button = ttk.Button(self.root, text = "Update", command = self.update_data)
        self.update_button.grid(row = 4, column = 5 , columnspan =2, padx = 10 , pady = 10)
#search button
        self.search_button = ttk.Button(self.root, text = "Search", command = self.search_data)
        self.search_button.grid(row = 3, column = 10 , columnspan =2, padx = 10 , pady = 10)
        
#search_data function (after click)
    def search_data(self):
        global df
        self.product_list.delete(0,tk.END)
        search_value = self.productToShow_entry.get()
        if (search_value == ''):
            messagebox.showerror("Error", "Enter product name to serach")
            return
        search_result = search_product(search_value,df)
        if (search_result.size != 0):
            height = search_result.size // 6
            rowspan = height
            window_width = "1250"
            window_height = "370"
            window_height = int(window_height) +  10*height
            window_height = str(window_height)
            self.root.geometry(f"{window_width}x{window_height}")
            self.product_list.grid_configure(rowspan=rowspan)
            self.product_list.config(height=height)
            for index, row in search_result.iterrows():
                entry_data = {
                'Product': row['Product'],
                'Quantity': row['Quantity'],
                'Location': row['Location'],
                'SerialNumber': row['SerialNumber'],
                'AirForceSerial': row['AirForceSerial'],
                'catalog': row['catalog']}
                self.product_list.insert(tk.END, entry_data)                     
        
        else:
            messagebox.showerror("Error", f"Product '{search_value}', Not found")
        self.productToShow_entry.delete(0,tk.END)
        
#update_quantity  function (after click)      
    def update_data(self):
        global df
        new_quantity = self.new_quantity_entry.get()
        remove_quantity = self.remove_quantity_entry.get()
        existProduct = self.update_entry.get()        
        if (existProduct == ''):
            messagebox.showerror("Error", "Enter product name")
            return
        if (new_quantity == '' and remove_quantity == ''):
            messagebox.showerror("Error", "Enter Quantity")
            return
        if (new_quantity.isdigit() == True):
            message = update_quantity(df,existProduct,new_quantity)
            messagebox.showinfo(message=message)
            self.new_quantity_entry.delete(0,tk.END)
            self.update_entry.delete(0,tk.END)
        elif(remove_quantity.isdigit() == True):
            remove_quantity = int(remove_quantity)*-1
            message = update_quantity(df,existProduct,remove_quantity)
            messagebox.showinfo(message=message)
            self.update_entry.delete(0,tk.END)
            self.new_quantity_entry.delete(0,tk.END)
            self.remove_quantity_entry.delete(0,tk.END)
            
            
#function to add add_data after click

    def submit_data(self):
        global df
        product = self.product_entry.get()
        quantity = self.quantity_entry.get()
        location = self.location_entry.get()
        serial_number = self.serial_entry.get()
        airforce_serial = self.airforce_entry.get()
        catalog = self.catalog_entry.get()
        
        if (product == ''):
            messagebox.showerror("No Product Added", "Enter Product Name")
            return
        if (quantity == ''):
            messagebox.showerror("No Quantity has been Added", "Enter Product Quantity")
            return
        if (quantity.isdigit() == False or int(quantity) == 0 ):
            messagebox.showerror("Number not Fit", "Quantity must be 1 and above")
            return
        if (location == ''):
            messagebox.showerror("No Location Added", "Enter Location")
            return
        
        entry_data = {'Product' : product,
                   'Quantity': quantity,
                   'Location': location,
                   'SerialNumber' : serial_number,
                   'AirForceSerial' : airforce_serial,
                   'catalog' : catalog}
        message = add_item(df,serial_number,product,quantity,location,airforce_serial,catalog)
        messagebox.showinfo(message=message)
        self.product_entry.delete(0,tk.END)
        self.quantity_entry.delete(0,tk.END)
        self.location_entry.delete(0,tk.END)
        self.serial_entry.delete(0,tk.END)
        self.airforce_entry.delete(0,tk.END)
        self.catalog_entry.delete(0,tk.END)
if __name__ == '__main__':
    app = App()
    app.root.mainloop()
                              


# In[35]:


# import pandas as pd
# columns = ['Product', 'Quantity', 'Location', 'SerialNumber', 'AirForceSerial', 'catalog']
# df = pd.DataFrame(columns=columns)
# excel_path = 'warehouse.xlsx'
# df.to_excel(excel_path,index=False)


# In[ ]:




