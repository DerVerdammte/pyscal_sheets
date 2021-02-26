# pyscal_sheets
Library for Controlling Google Sheets with Pascals *legendary* unusable code!

Funktionen:
pyscal_sheets(id) Erstellt ein Token, sofern es noch nicht da ist, und erlaubt das benutzen der anderen Befehle.
sheet = pyscal_sheets(id)

sheet.fill_range(range, values)
    returns None
    
sheet.delete_range(range)
    returns delted_values
    
sheet.update_range(range, value)
    returns None
    
sheet.update_field(field, value)
    returns None
    
sheet.append_row(range, row)
    returns None
    
sheet.read_range(range)
    returns values
    
sheet.read_field(field)
    returns value
    
sheet.search_for(range, keyword, column)
    returns row




