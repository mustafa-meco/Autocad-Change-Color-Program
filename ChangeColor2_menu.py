import comtypes.client
import tkinter as tk
import ttkbootstrap as ttk # import ttkbootstrap

# Connect to AutoCAD's COM interface
acad = comtypes.client.GetActiveObject("AutoCAD.Application")
# Get the active document
doc = acad.ActiveDocument

def send_command(group_name, color):
  command = f"(c:mycommand \"{group_name}\" {color})"
  doc.SendCommand(command+"\n")

def get_groups():
   groups = []
   for group in doc.Groups:
      groups.append(group.Name)
   return groups

def submit(event=None):
    group_name = group_list.get()
    color = color_entry.get()
    groups = get_groups()
    if group_name == "":
       result_label.config(text=f"Please enter group name!", bootstyle="danger")
       return
    elif color == "":
       result_label.config(text=f"Please enter color!", bootstyle="danger")
       return
    elif color.isdigit() == False:
       result_label.config(text=f"Please enter color as number!", bootstyle="danger")
       return
    elif group_name.upper() not in groups:
       result_label.config(text=f"Please enter correct group name!", bootstyle="danger")
       return
    
    send_command(group_name , color )
    result_label.config(text=f"Group Name: {group_name}\nColor: {color}", bootstyle="success" )

# Create the main window
window = ttk.Window(themename="superhero") # create a Window object with a theme
window.title("Change Group Color | AutoCAD / by M.Ghoneim")
window.geometry("550x250") # set window size

# Create and place Entry widgets for group name and color
group_name_label = ttk.Label(window, text="Group Name:", bootstyle="primary", font=('helvetica', 12, 'bold')) # use bootstyle parameter
group_name_label.pack()
# group_name_entry = ttk.Entry(window, bootstyle="primary")
# group_name_entry.pack()
# group_name_entry.focus() # set focus to this entry widget
# group_name_entry.bind("<Return>", submit) # bind the <Return> key to the submit function
# Make the group in a menu instead of a entry to choose from it, use the group list
group_list = ttk.Combobox(window, values=get_groups(), bootstyle="primary")
group_list.pack()
group_list.bind("<Return>", submit) # bind the <Return> key to the submit function



color_label = ttk.Label(window, text="Color:", bootstyle="primary", font=('helvetica', 12, 'bold'))
color_label.pack()
color_entry = ttk.Entry(window, bootstyle="primary")
color_entry.pack()
color_entry.bind("<Return>", submit) # bind the <Return> key to the submit function

# Create a Submit button
submit_button = ttk.Button(window, text="Submit", command=submit, bootstyle="success") # use bootstyle parameter
submit_button.pack(pady=10)

# Create a label to display the result
result_label = ttk.Label(window, text="", font=('helvetica', 12, 'bold'))
result_label.pack()

# Start the main loop
window.mainloop()
