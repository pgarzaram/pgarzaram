import os
import barcode
from barcode.writer import ImageWriter
import webbrowser

def generate_barcode(data, barcode_type='code128', output_folder='Barcode_Folder_Storage'):
    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Create a Barcode object
    code = barcode.get_barcode_class(barcode_type)

    # Set the data and create the barcode
    code_instance = code(data, writer=ImageWriter())

    # Save the barcode to an image file in the specified folder
    filename = f"{output_folder}/{data}_{barcode_type}"
    code_instance.save(filename)

    print(f"Barcode saved as {filename}")

# Main loop for generating barcodes
while True:
    data_to_encode = input("Value to barcode (Type 'quit' to exit): ")

    if not data_to_encode:
        print("Please enter a value.")
        continue
    elif data_to_encode.lower() == 'quit':
        print("Exiting the program.")
        break

    generate_barcode(data_to_encode, barcode_type='code128', output_folder='Barcode_Folder_Storage')


