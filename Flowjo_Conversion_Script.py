import argparse
import xlrd
from xlutils.copy import copy
import xlwt
import pandas as pd
import statistics

#God I wish I knew about Pandas before writing this
def Replace_FLowJo_Output(Input_Name, Output_Name):

    if Output_Name is None:
        Output_Name = Input_Name[:-4] + "_converted.xls"
    FlowJo = xlrd.open_workbook(Input_Name)
    FlowJo_Sheet = FlowJo.sheet_by_index(0)

    Sheet_Name = "FlowJo Output"

    #Test to see if the a output file already exists
    File_Exists = True
    Previous_Position = 0
    try:
        Output_File = xlrd.open_workbook(Output_Name)
    except FileNotFoundError:
        Excel_Output = xlwt.Workbook()
        Output_Sheet = Excel_Output.add_sheet(Sheet_Name)
        File_Exists = False

    if(File_Exists):
        Output_Sheet = Output_File.sheet_by_index(Output_File.nsheets - 1)
        Previous_Position = Output_Sheet.nrows + 2
        Excel_Output = copy(Output_File)
        Output_Sheet = Excel_Output.get_sheet(Output_File.nsheets - 1)

    Output_Sheet.write(Previous_Position + 0,0, "Run Name")
    Output_Sheet.write(Previous_Position + 0,1, "Plate Number")
    Output_Sheet.write(Previous_Position + 0,2, "Well Label")

    Replicates = []
    index = 1
    for i, Run_name in enumerate(FlowJo_Sheet.col_values(0)):
        if (FlowJo_Sheet.cell(i+1,0).value[3:FlowJo_Sheet.cell(i+1,0).value.rfind("-")-len(FlowJo_Sheet.cell(i+1,0).value)-2] == Run_name[3:Run_name.rfind("-")-len(Run_name)]):
            Replicates.append(True)
        else:
            Replicates.append(False)
        if i == 0 or Replicates[i-1]:
            continue
        Output_Sheet.write(Previous_Position + index, 0, Run_name[3:Run_name.rfind("-")-len(Run_name)])
        Output_Sheet.write(Previous_Position + index, 1, Run_name[:2])
        Output_Sheet.write(Previous_Position + index, 2, Run_name[Run_name.rfind("-")-len(Run_name)+1:-4])
        if (FlowJo_Sheet.cell(i+1,0).value == "Mean"):
            Output_Sheet.write(Previous_Position + index + 1, 0, FlowJo_Sheet.cell(i+1,0).value)
            Output_Sheet.write(Previous_Position + index + 2, 0, FlowJo_Sheet.cell(i+2,0).value)
            Replicates.extend([False,False])
            break
        index += 1

    for i in range(FlowJo_Sheet.ncols - 1):
        index = 1
        for j, Cell in enumerate(FlowJo_Sheet.col_values(i+1)):
            if j == 0:
                Output_Sheet.write(Previous_Position + j, i + 3, Cell)
                continue
            if Replicates[j-1]:
                continue
            if isinstance(Cell, str) and Cell.endswith(" %"):
                Cell = Cell[:-2]
                Cell = Cell.replace(",", ".")
            Values = []
            Values.append(float(Cell))
            k = 0
            while Replicates[j+k]:
                Next_Cell = FlowJo_Sheet.cell_value(j+k+1, i+1)
                if isinstance(Next_Cell, str) and Next_Cell.endswith(" %"):
                    Next_Cell = Next_Cell[:-2]
                    Next_Cell = Next_Cell.replace(",", ".")
                Values.append(float(Next_Cell))
                k += 1
            Cell_Value = statistics.mean(Values)
            Output_Sheet.write(Previous_Position + index, i + 3, Cell_Value)
            index += 1

    Excel_Output.save(Output_Name)

    #New Pandas code, it works and this isn't that urgent code
    Results = pd.read_excel(Output_Name)

    Temp_Frame = Results[["Plate Number","Well Label"]]
    Temp_Frame["Well Label"] = pd.to_numeric(Temp_Frame["Well Label"].str.replace('[a-zA-Z]', ''), errors="coerce")
    Temp_Frame = Temp_Frame.sort_values(by = ["Plate Number","Well Label"])

    Results.reindex(Temp_Frame.index).to_excel(Output_Name)

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Get FlowJo Files to convert into Excel Output")
    parser.add_argument( "Input_Files", nargs='+', help='Path to the FlowJo Excel File ')
    parser.add_argument("-O", "--Output_File", metavar = "", help= """Desired Path to the Output Excel File or Path
                                                          to the current Excel File you would like to append data to""")

    args = parser.parse_args()
    for File in args.Input_Files:
        Replace_FLowJo_Output(File, args.Output_File)
