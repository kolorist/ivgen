import openpyxl
import xml.dom.minidom


def entry_point():
    print("invoice_generator v0")

    doc = xml.dom.minidom.parse("849804.xml")
    hdon = doc.getElementsByTagName("HDon")[0]
    dlhdon = hdon.getElementsByTagName("DLHDon")[0]
    ndhdon = dlhdon.getElementsByTagName("NDHDon")[0]
    dshhdvu = ndhdon.getElementsByTagName("DSHHDVu")[0]
    hhdvuArr = dshhdvu.getElementsByTagName("HHDVu")

    for hhdvu in hhdvuArr:
        mhhdvu = hhdvu.getElementsByTagName("MHHDVu")[0]
        print(mhhdvu.firstChild.data)

    wb = openpyxl.load_workbook("template.xlsx")
    print(wb.sheetnames)

    ws = wb['Sheet1']
    c0 = ws['A3']
    print(c0.value)
    ws['A5'] = "1\n2\n3"
    ws['A5'].alignment = openpyxl.styles.Alignment(wrapText=True)

    ws.insert_rows(5, 3)

    wb.save("output.xlsx")


if __name__ == "__main__":
    entry_point()
