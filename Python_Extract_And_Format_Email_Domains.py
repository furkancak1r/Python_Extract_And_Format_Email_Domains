# Mail adreslerindeki .com veya .com.tr'ye kadar olan kısmı alan kod
import openpyxl

def extract_domain(cell_value):
    if ".com.tr" in cell_value:
        return cell_value.split(".com.tr")[0] + ".com.tr"
    elif ".com" in cell_value:
        return cell_value.split(".com")[0] + ".com"
    else:
        return cell_value


# Excel dosyasını aç
wb = openpyxl.load_workbook(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")

sheet = wb['INKOOL']  # Sayfa adını buraya yazın

# J5'ten J7583'e kadar olan hücreleri dolaş
for row in range(5, 7584):
    cell_value = sheet[f'J{row}'].value

    # Eğer hücre boş değilse, domain'i çıkar ve K sütununa yaz
    if cell_value is not None:
        extracted_value = extract_domain(cell_value)
        small_letters = extracted_value.lower()
        sheet[f'L{row}'].value = small_letters

# Excel dosyasını kaydet
wb.save(r"C:\Users\furkan.cakir\Desktop\FurkanPRS\Kodlar\Finans & Muhasebe\exceller\Cariler\OCPR-İlgili Kişiler.xlsx")
