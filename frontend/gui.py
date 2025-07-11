import os
import tkinter as tk
from tkinter import filedialog, messagebox

from backend.image_to_ppt import create_ppt_from_image


def select_image():
    path = filedialog.askopenfilename(filetypes=[('Image Files', '*.png *.jpg *.jpeg *.bmp *.gif')])
    if path:
        image_var.set(path)


def select_output():
    path = filedialog.asksaveasfilename(defaultextension='.pptx', filetypes=[('PowerPoint', '*.pptx')])
    if path:
        output_var.set(path)


def convert():
    img = image_var.get()
    out = output_var.get()
    if not img or not os.path.isfile(img):
        messagebox.showerror('Error', 'Please choose a valid image file.')
        return
    if not out:
        out = os.path.splitext(img)[0] + '.pptx'
    try:
        create_ppt_from_image(img, out)
        messagebox.showinfo('Success', f'Saved PPT to {out}')
    except Exception as e:
        messagebox.showerror('Error', str(e))


root = tk.Tk()
root.title('Image to PPT Converter')

image_var = tk.StringVar()
output_var = tk.StringVar()

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

# Image selection
img_label = tk.Label(frame, text='Image:')
img_label.grid(row=0, column=0, sticky='e')
img_entry = tk.Entry(frame, textvariable=image_var, width=40)
img_entry.grid(row=0, column=1)
img_button = tk.Button(frame, text='Browse...', command=select_image)
img_button.grid(row=0, column=2, padx=5)

# Output selection
out_label = tk.Label(frame, text='Output PPT:')
out_label.grid(row=1, column=0, sticky='e')
out_entry = tk.Entry(frame, textvariable=output_var, width=40)
out_entry.grid(row=1, column=1)
out_button = tk.Button(frame, text='Save As...', command=select_output)
out_button.grid(row=1, column=2, padx=5)

convert_button = tk.Button(frame, text='Convert', command=convert)
convert_button.grid(row=2, column=1, pady=10)

root.mainloop()
