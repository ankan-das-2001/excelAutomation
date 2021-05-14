import os
from openpyxl import Workbook
import pandas as pd
import csv
import numpy
import xlsxwriter
from IPython.display import HTML

Dataframes=[]
Competitor_name = []
alphabets = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']


def CreatingCompetitorTable(new_df,competitor_name):
    competitor_array = new_df[competitor_name]
    URL_array=new_df['URL']

    data=[]
    index =0            # initialised index here
    hyperlink_list=[]
    for i in range(len(URL_array)):
        data.append('=HYPERLINK("{}","{}")'.format(URL_array[i],competitor_array[i]))
    new_df = pd.DataFrame({ competitor_name : data})
    new_df.to_excel("test.xlsx",index = False, header=True)
    return new_df

    # close the workbook


def parseWorkBook(previous_dirnames,file):
    if file=="11. Keywords info.csv":               #checking whether the file with name 11. Keywords info.csv exist or not
        if previous_dirnames=="" or len(previous_dirnames)==0:             #checking the directory name is parent or not
            df=pd.read_csv(cwd+"\\"+file,encoding="UTF-16",sep='\t')
        else:

            df=pd.read_csv(previous_dirnames+"\\"+file,encoding="UTF-16",sep='\t')

        # droping unneccesary collumns from our dataset
        new_df = df.drop(['#','Position History','Position History Date','CPC','Last Update','Page URL inside','SERP Features'],axis=1)

        # This bunch of code is actually used to assign the value competitor names
        keyword_array = new_df['Keyword'].to_numpy()
        s = previous_dirnames.rsplit('/',1)
        competitor_name =s[1]

        #Adding the competitor
        Position = new_df['Position']
        Traffic = new_df['Traffic (desc)']
        competitor = []
        for i in range(len(Position)):
            p = Position[i]
            t = Traffic[i]
            competitor.append("{}/{}".format(str(p),str(t)))


        new_df[competitor_name] = competitor
        hyperlink_df = CreatingCompetitorTable(new_df, competitor_name)

        Competitor_name.append(competitor_name)
        new_df[competitor_name] = hyperlink_df[competitor_name]
        new_df=new_df.drop(['Position','URL'],axis=1)

        #Final
        Dataframes.append(new_df)



def parsingLast():
    # Concatanating all the dataframes into one common dataframe named new_dataframe
    new_dataframe = pd.concat(Dataframes)

    # Sorting the dataframe calculating the


    #n_df = new_dataframe.drop(['Position','URL','Traffic (desc)'],axis=1)

    dictionary_duplicate_remover = dict()
    #dictionary_duplicate_remover['Keyword'] = 'first'
    dictionary_duplicate_remover['Volume'] = 'first'
    dictionary_duplicate_remover['Difficulty'] = 'first'
    dictionary_duplicate_remover['Traffic (desc)'] = 'first'
    for names in Competitor_name:
        dictionary_duplicate_remover[names] = 'first'



    # Now removing the data previously duplicated
    n_df = new_dataframe.groupby('Keyword').agg(dictionary_duplicate_remover).reset_index()

    n_df.sort_values(by=['Volume'],inplace=True, ascending=False)
    print(n_df.head(5))
    n_df = n_df.drop(['Traffic (desc)'],axis=1)
    n_df.style.set_properties(**{'text-align':'center'})
    n_df.to_excel("Competitor_name.xlsx",sheet_name= "Competitor_data",index = False, header=True)

    return n_df


def decorating_excel_sheet(df):
    writer = pd.ExcelWriter('Competitor_name.xlsx')



    df.to_excel(writer, sheet_name='Competitor_data', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Competitor_data']

    header_format = workbook.add_format()
    header_format.set_bold()
    header_format.set_center_across()

    normal_format= workbook.add_format()
    normal_format.set_align('center')

    link_format = workbook.add_format()
    link_format.set_font_color('blue')
    link_format.set_align('center')
    link_format.set_underline()
    index=0
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        if index==0:
            writer.sheets['Competitor_data'].set_column(col_idx, col_idx, 40,normal_format)
        elif index>0 and index<3:
            writer.sheets['Competitor_data'].set_column(col_idx, col_idx, 15,normal_format)
        elif index>=3:
            writer.sheets['Competitor_data'].set_column(col_idx, col_idx, 20,link_format)
        index+=1
        #writer.sheets['Competitor_data'].add_format().set_align('center')
    writer.save()

print("Enter the location of parent folder where you have information about the subfolder: ")
filepath = input()
print("Wait while we collect the data: ")

for dirpath, dirnames, filenames in os.walk(filepath):
    # for terminal cmd usage printing all the files or folder names.
    #print("directory-path: {}".format(dirpath))
    #print("directory name: {}".format(dirnames))
    #print("File names: {}".format(filenames))

    previous_dirnames =  dirnames
    # After processing the file directories, we will parse each and every file structures.
    for file in filenames:
        parseWorkBook(dirpath,file)
n_df = parsingLast()
decorating_excel_sheet(n_df)
