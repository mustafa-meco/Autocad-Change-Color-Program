import comtypes.client
import tkinter as tk


def send_command(group_name, color):
  # Connect to AutoCAD's COM interface
  acad = comtypes.client.GetActiveObject("AutoCAD.Application")
  # Get the active document
  doc = acad.ActiveDocument

  command = f"(c:mycommand \"{group_name}\" {color})"
  doc.SendCommand(command+"\n")




def submit():
    group_name = group_name_entry.get()
    color = color_entry.get()
    send_command(group_name , color )
    result_label.config(text=f"group Name: {group_name}\nColor: {color}")

# Create the main window
window = tk.Tk()
window.title("group Info")

# Create and place Entry widgets for group name and color
group_name_label = tk.Label(window, text="group Name:")
group_name_label.pack()
group_name_entry = tk.Entry(window)
group_name_entry.pack()

color_label = tk.Label(window, text="Color:")
color_label.pack()
color_entry = tk.Entry(window)
color_entry.pack()

# Create a Submit button
submit_button = tk.Button(window, text="Submit", command=submit)
submit_button.pack()

# Create a label to display the result
result_label = tk.Label(window, text="")
result_label.pack()

# Start the main loop
window.mainloop()