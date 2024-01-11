import xlsxwriter
from os.path import expanduser

while True:
    try:
        desktop = expanduser("~/Desktop")

        name = input('name of file : ')
        dastmozd = float(input('dastmozd mabna : '))
        year = int(input('starter year (1396,...,1400) : '))

        if year == 1396:
            S_year = 0
        elif year == 1397:
            S_year = 1
        elif year == 1398:
            S_year = 2
        elif year == 1399:
            S_year = 3
        elif year == 1400:
            S_year = 4
        else:
            S_year = "error"

        data = []
        members = []
        year_reader = 1396
        month_reader = 1

        E96 = (1.2 * dastmozd) + 6768
        E97 = (1.104 * E96) + 28208
        E98 = (1.13 * E97) + 87049
        E991 = (1.15 * E98) + 30338
        E994 = (1.15 * E98) + 50338
        E00 = (1.26 * E994) + 82785

        E_List = [E96, E97, E98, E991, E00]
        E_List = E_List[S_year:]

        for item in E_List:
            if item == E991:
                while month_reader <= 4:
                    members = [year, month_reader, int(item * 31)]
                    data.append(members)
                    month_reader += 1
                while month_reader <= 6:
                    members = [year, month_reader, int(E994 * 31)]
                    data.append(members)
                    month_reader += 1
                while month_reader <= 12:
                    members = [year, month_reader, int(E994 * 30)]
                    data.append(members)
                    month_reader += 1
                month_reader = 1
                year += 1
            else:
                while month_reader <= 6:
                    members = [year, month_reader, int(item * 31)]
                    data.append(members)
                    month_reader += 1
                while month_reader < 12:
                    members = [year, month_reader, int(item * 30)]
                    data.append(members)
                    month_reader += 1
                while month_reader == 12:
                    members = [year, month_reader, int(item * 29)]
                    data.append(members)
                    month_reader += 1
                month_reader = 1
                year += 1

        workbook = xlsxwriter.Workbook(f'{desktop}\\{name}.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'year')
        worksheet.write('B1', 'month')
        worksheet.write('C1', 'eslahi')

        row = 1
        col = 0

        for S, M, D in data:
            worksheet.write(row, col, S)
            worksheet.write(row, col + 1, M)
            worksheet.write(row, col + 2, D)
            row += 1

        workbook.close()
        print("\n done")
        print("--------------------- \n")
    except:
        print("an ERROR accured \n you may type a wrong year \n Or there is a bug in program! \n please try again! ")
