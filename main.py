from ExcelManipulator import ExcelManipulator

if __name__ == '__main__':

    em = ExcelManipulator("example.xlsx")
    em.workFunction([2, 3, 'D', 'E'])
    em.saveWB()

