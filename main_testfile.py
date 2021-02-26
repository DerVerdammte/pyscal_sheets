from Pyscal import Pyscalsheets

pyscal = Pyscalsheets("1ZhVw2du5qQ_oBQdTN4FXLfDZCgR95o7IbCGDstekrCc",
                      "1NG-Avb1WymSAfApRnCK6BIwevV_ssn187WqEbvFLU7c",
                      1) ##Pyscal
# Sheets
print(f"First  Spreadsheet Id: {pyscal.get_spreadsheet_id()}")
print(f"Second Spreadsheet Id: {pyscal.get_log_spreadsheet_id()}")

pyscal.append_spreadsheet_row(pyscal.generate_body([1,2,3,4],
                                                   insert_time=True))
#pyscal.get_data_from_sheet()

#print(pyscal.find_unique("Pascal", 2))

#print(pyscal.find_all_keys("Pascal", 2))
#print("Vallah")
#print(pyscal.convert_to_range(0,0,1,3))

#pyscal.replace_range(1,3, [[1,2,3],[4,5,6],[7,8,9]])
#pyscal.delete_coords(1,3,2,4)
#pyscal.delete_coords(9,10,2,3)
pyscal.replace_field(1,1, ["Pimmel"])

pyscal.replace_field(0,3, ["Schwanz"])

pyscal.replace_field(5,0, ["Rödel"])


pyscal.replace_field(5,6, ["Wubbel"])