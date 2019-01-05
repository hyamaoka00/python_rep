import urllib3
from bs4 import BeautifulSoup


def edit_excel(sheet, url, url_add, i, j, k):

    for cell_obj in list(sheet.columns)[j]:

        if cell_obj.value is None:
            break

        if str(cell_obj.value) != "銘柄":

            if str(cell_obj.value) == "RDSB":
                    cell_obj.value = "RDS/B"

            # url組み立て
            new_url = url + cell_obj.value + url_add

            # urllib3を使うならこっち
            http = urllib3.PoolManager()
            r = http.request("GET", new_url)

            soup = BeautifulSoup(r.data, "html.parser")
            search_div_tag = soup.find_all("div")

            price = ""

            for tag in search_div_tag:
                try:
                    string_ = tag.get("class").pop(0)
                    if string_ in "price":
                        price = tag.string
                        break
                except:
                    pass

            sheet.cell(row=i, column=k, value=price)

            i = i + 1

    return sheet
