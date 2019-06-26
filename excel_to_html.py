import sys
from pathlib import Path
import xlrd

p = Path('.')


def get_input_output_file(args) :
    if not args:
        sys.exit(0)
    else:
        if len(args) != 3:
            print("usage is excel_to_html input.xls output.html")
        else:
            if '\\' in args[1] :
                args[1].replace('\\', '/')
                args[2].replace('\\', '/')
            path_to_input = str(args[1]).split('/')
            input = p
            for folder in path_to_input :
                input = input / folder
            path_to_output = args[2].split('/')
            output = p
            for folder in path_to_output :
                output = output / folder
            print(input, output)
            return input, output


def convert_xls_to_html(xls_file):
    # adding css
    html = "<style>\n"
    html += ".tab { \n" \
            " overflow: hidden; " \
            " border: 1px solid #ccc; " \
            " background-color: #f1f1f1; }\n" \
            ".tab button {\n" \
            " background-color: inherit;" \
            " float: left; " \
            " border: none;" \
            " outline: none; " \
            " cursor: pointer; " \
            " padding: 14px 16px;" \
            " transition: 0.3s; }\n" \
            ".tab button:hover {\n " \
            " background-color: #ddd; }\n" \
            ".tab button.active {" \
            " background-color: #ccc;}\n" \
            ".tabcontent {" \
            " display: none;" \
            " padding: 6px 12px;" \
            " border: 1px solid #ccc; " \
            " border-top: none; }\n"
    html += "</style>\n"
    html += "<div class=\"tab\">\n"
    html += "<button class=\"tablinks\" onclick=\"changeSheet(event, \'" + str(xls_file.sheet_by_index(0).name).replace("'","") + "\')\" id=\"defaultOpen\">" + xls_file.sheet_by_index(0).name + "</button>\n"
    for i in range( 1, xls_file.nsheets):
        html += "<button class=\"tablinks\" onclick=\"changeSheet(event, \'" + str(xls_file.sheet_by_index(i).name).replace("'", "") +"\')\">" + xls_file.sheet_by_index(i).name + "</button>\n"
    html += "</div>\n"
    for sheet in xls_file.sheets():
        html += "<div id=\"" + str(sheet.name).replace("'", "")+ "\" class=\"tabcontent\">\n"
        html += "<table class=\"table table-dark\">\n<tr>\n"
        for cell in sheet.row(0):
            html+= "<th scope=\"col\">" + cell.value + "</th>\n"
        #print(html)
        html += "</tr>\n"
        for i in range(1, sheet.nrows):
            print(sheet.row(i))
            for cell in sheet.row(i):
                html += "<td scope=\"row\">" + str(cell.value) + "</td>\n"
            #print(html)
            html += "</tr>\n"
        html += "</table>\n</div>\n\n"

        #adding js for buton
    html += "<script type=\"text/javascript\">\n"
    html += "function changeSheet(evt, sheetName){ \n " \
            "  var i, tabcontent, tablinks;\n " \
            "  tabcontent = document.getElementsByClassName(\"tabcontent\");\n" \
            "  for (i = 0; i < tabcontent.length; i++) { " \
            "  tabcontent[i].style.display = \"none\";}\n" \
            "  tablinks = document.getElementsByClassName(\"tablinks\");\n" \
            "  for (i = 0; i < tablinks.length; i++) { " \
            "  tablinks[i].className = tablinks[i].className.replace(\"active\", \"\");}\n" \
            "  document.getElementById(sheetName).style.display = \"block\";\n" \
            "  evt.currentTarget.className +=  \" active\"; }" \
            "  document.getElementById(\"defaultOpen\").click();\n"
    html += "</script>\n"

    return html


def main():
    print(sys.argv)
    input, output = get_input_output_file(sys.argv)
    xls_file = xlrd.open_workbook(input)
    html_data = convert_xls_to_html(xls_file)
    with open(output, "w") as out :
        out.write(html_data)


if __name__ == "__main__":
    main()