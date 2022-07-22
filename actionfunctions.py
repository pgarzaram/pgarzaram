def CheckOut(max_row, ws):
    Status = "repeat"
    while Status == "repeat":
        print("CHECK OUT ITEMS")
        CheckOut_Name = str(input("Input Part Number: "))
        CheckOut_Num = int(input("Quantity: "))
        print("Please wait....\n")

        for j in range(1, max_row + 1):
            column_1 = ws.cell(row=j, column=1)
            if column_1.value == CheckOut_Name:
                print(f"Part Number: {column_1.value}")
                chang_val = ws.cell(row=j, column=3)
                print(f"Original In Stock Number: {chang_val.value}")
                chang_val.value = chang_val.value - CheckOut_Num
                print(f"New In Stock Number: {chang_val.value}")
                Status = "PASS"
                break
        if Status == "repeat":
            print("You're part number was not found, please try again.\n")


def CheckIn(max_row, ws):
    Status = "repeat"
    while Status == "repeat":
        print("CHECK IN ITEMS")
        CheckIn_Name = str(input("Input Part Number: "))
        CheckIn_Num = int(input("Quantity: "))
        print("Please wait....\n")

        for j in range(1, max_row + 1):
            column_1 = ws.cell(row=j, column=1)

            if column_1.value == CheckIn_Name:
                print(f"Part Number: {column_1.value}")
                chang_val = ws.cell(row=j, column=3)
                print(f"Original In Stock Number: {chang_val.value}")
                chang_val.value = chang_val.value + CheckIn_Num
                print(f"New In Stock Number: {chang_val.value}")
                Status = "PASS"
                break
        if Status == "repeat":
            print("You're part number was not found, please try again.\n")
