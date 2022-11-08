import mysql.connector
import pandas as pd
import openpyxl



def access():
    while(True):
        user_name = input("user name >> ")
        password =  input("password  >> ")
        database = input("Database  >> ")

        try:
            local = mysql.connector.connect(
            host = "localhost",
            port = 3306,
            user = user_name,
            password = password,
            database = database
            )
            break
    
        except:
            print("데이터 베이스 읽기 오류입니다.")

    print("\n데이터 베이스 접속을 환영합니다. \n 유저 : " 
        + user_name +
        "\n데이터베이스 : " + database )
        
    return local







def make_table():
    table_type_list = []
    table_name_list = []

    print("테이블을 제작합니다. \n ")
    
    print("excel, csv파일을 참고하여 작성하시겠습니까? \n")
    sel_csv = int(input("yes : 1번 , no: 2번 >> "))
    if sel_csv == 2:
        table_name = input("테이블 이름을 입력하세요. >>")
        table_size = int(input("컬럼의 개수를 입력하세요. >>"))

        for i in range(table_size):
            print(str(i+1) + "번째 데이터를 생성합니다. \n")
            table_type_list.append(input("컬럼 타입을 입력하세요. >>"))
            table_name_list.append(input("컬럼 이름을 입력하세요. >>"))
            print("테이블 지정 완료 \n")
    
        sql = "create table " + table_name + "\n" + \
         "("
    
        for i in range(table_size):
            sql = sql +" `" + table_name_list[i] + "` " + table_type_list[i]
            if(i != table_size-1):
                sql = sql + "," + "\n"
     
        sql = sql + " );"
        print("명령어를 확인하세요. \n"  + sql)
        sel = int(input("다음의 명령어를 삽입하시겠습니까? 1을 입력하면 됩니다."))
        if(sel == 1):
            return sql
        else:
            return 1





    elif sel_csv == 1:
        try:
            sel_file = int(input("excel은 1번 csv는 2번 \n"))
            if(sel_file == 1):
                try:
                    excel_file = input("excel파일 이름을 입력, \n주의! 같은 폴더에 있을 것 \n주의! .xlsx까지 입력할 것\n>>")
                    df = pd.DataFrame(pd.read_excel("./" + excel_file))
                except:
                    print("인코딩 문제가 발생했습니다. EUC-KR로 적용합니다.")
                    file = "./" + excel_file
                    df = pd.DataFrame(pd.read_excel(file, encoding='EUC-KR'))

            if(sel_file == 2):
                try:
                    csv_file = input("csv파일 이름을 입력, \n주의! 같은 폴더에 있을 것 \n주의! .csv까지 입력할 것\n>>")
                    df = pd.DataFrame(pd.read_csv("./" + csv_file))
                except:
                    print("인코딩 문제가 발생했습니다. EUC-KR로 적용합니다.")
                    file = "./" + csv_file
                    df = pd.DataFrame(pd.read_csv(file, encoding='EUC-KR'))
        



            #df_list_tmp = df.columns.values.tolist()
            df_list = df.columns.values.tolist()
            #for i in df_list:
            #   temp = i.replace(" ", ""_)

            table_name = input("테이블 이름을 입력하세요. >>")
            table_type_list = []

            auto_type = int(input("데이터 형태를 excel, csv에 따라만들겠습니까? \n시간이 단축되나 날짜형을 지정할 수 없습니다. yes = 1 , no = 2 >>"))

            if auto_type == 2:
                for i in range(len(df_list)):
                    print(str(len(df_list)) + "개의 데이터를 입력해야합니다. \n")
                    print(str(i+1) + "번째 데이터를 생성합니다. \n")
                    table_type_list.append(input(df_list[i] + "의 변수타입을 입력하세요. >>"))
                    print("테이블 지정 완료 \n")
    
                sql = "create table " + table_name + "\n" + \
                "("
    
                for i in range(len(df_list)):
                    sql = sql +" `" + df_list[i] + "` " + table_type_list[i]
                    if(i != len(df_list)-1):
                        sql = sql + "," + "\n"
     
                sql = sql + " );"
                print("명령어를 확인하세요. \n"  + sql)
                sel = int(input("다음의 명령어를 삽입하시겠습니까? 1을 입력하면 됩니다."))
                if(sel == 1):
                    return sql
                else:
                    return 1



            if auto_type == 1:
                print("csv파일의 타입을 따라 만듭니다.")
                df_type = pd.DataFrame(
                        {'int' : ['1'],
                         'float' : ['1.0'],
                         'date' : ['2022-11-07'],
                         'object' : ["string"]
                       }) 
                df_type = df_type.astype({'int':'int'})
                df_type = df_type.astype({'float':'float'})
        
                for each in range(len(df_list)):
                    for i in range(len(df_type)):
                        if df[df_list[each]].dtype == df_type['int'].dtype:
                            table_type_list.append("int")
                        elif df[df_list[each]].dtype == df_type['float'].dtype:
                            table_type_list.append("float")
                        elif df[df_list[each]].dtype == df_type['object'].dtype:
                            table_type_list.append("varchar(100)")
    
                sql = "create table " + table_name + "\n" + \
                "("
    
                for i in range(len(df_list)):
                    sql = sql +" `" + df_list[i] + "` " + table_type_list[i]
                    if(i != len(df_list)-1):
                        sql = sql + "," + "\n"
     
                sql = sql + " );"
                print("명령어를 확인하세요. \n"  + sql)
                sel = int(input("다음의 명령어를 삽입하시겠습니까? 1을 입력하면 됩니다."))
                if(sel == 1):
                    return sql
                else:
                    return 1


        except:
            print("파일을 읽는데 실패했습니다.")




def show_table(cursor):
    print("현재 생성된 테이블 목록입니다. \n")
    sql = "show tables"
    cursor.execute(sql)

    result = cursor.fetchall()
    for resul_iterator in result:
        print(resul_iterator)
    print("\n")








def delete_table(cursor):
        show_table(cursor)
        print("어떤 테이블을 삭제할까요? \n")
        table_name = input(">> ")
        sql = "drop table " + table_name 
        print("주의!!! 조원의 동의를 구했나요? \n백업을 했나요?")
        real = int(input("진짜? 삭제를 원하면 1을 입력하세요."))
        
        if(real == 1):
            try:
                cursor.execute(sql)
                print("삭제를 완료했습니다.\n\n")
            except:
                print("테이블 이름 오류입니다. 메인메뉴로 이동합니다.")

        else:
            print("데이터는 소중합니다\n\n.")



def insert_data(cursor, local):
        show_table(cursor)
        print("데이터를 삽입합니다.")
        print("어떤 테이블에 삽입할까요? \n")
        table_name = input(">> ")
        try:
            sql = "desc " + table_name
            cursor.execute(sql)

            result = cursor.fetchall()
            for resul_iterator in result:
                print(resul_iterator)
            print("\n")

            print("데이터 삽입 주의 사항! \n테이블의 컬럼수와 csv, excel의 컬럼수가 일치해야합니다.\n")
            
            sel_file = int(input("excel(xlsx)은 1번 csv는 2번 \n"))
            if(sel_file == 1):
                try:
                    excel_file = input("excel파일 이름을 입력, \n주의! 같은 폴더에 있을 것 \n주의! .xlsx까지 입력할 것\n>>")
                    df = pd.DataFrame(pd.read_excel("./" + excel_file))
                except:
                    print("인코딩 문제가 발생했습니다. EUC-KR로 적용합니다.")
                    file = "./" + excel_file
                    df = pd.DataFrame(pd.read_excel(file, encoding='EUC-KR'))

            if(sel_file == 2):
                try:
                    csv_file = input("csv파일 이름을 입력, \n주의! 같은 폴더에 있을 것 \n주의! .csv까지 입력할 것\n>>")
                    df = pd.DataFrame(pd.read_csv("./" + csv_file))
                except:
                    print("인코딩 문제가 발생했습니다. EUC-KR로 적용합니다.")
                    file = "./" + csv_file
                    df = pd.DataFrame(pd.read_csv(file, encoding='EUC-KR'))

            x_count = df.shape[1]
            sql = "insert into " + table_name + " values ("
            for i in range(x_count):
                sql = sql + " %s"
                if(i != x_count -1):
                    sql = sql + ", "
            sql = sql + " );"
            print(sql)

            count = 0
            max = 1000
            print(str(len(df)) + "개의 데이터를 삽입합니다.")

            for i, row in df.iterrows():
                cursor.execute(sql, tuple(row))
                local.commit()
                count = count +1
                if count >= max:
                    print(str(count) + " / " + str(len(df))  )
                    max = max + 1000
                
            print("데이터 삽입 종료!! \n\n")

        except:
            print("오류가 발생했습니다. 처음으로 돌아갑니다.\n\n")

#main
local = access()
cursor = local.cursor(buffered=True)


while(True):
    print(
        "작업을 선택하세요. \n"
        + "1. 테이블 제작 \n"
        + "2. 데이터 삽입 \n"
        + "3. 테이블 삭제 \n" 
        + "4. 조회 \n"
        + "5. 종료 \n"
      )
    sel = input(">>")

    if int(sel) == 1:
        sql = make_table()
        try:
            if(sql != 1):
                cursor.execute(sql)
                print("테이블 제작 성공\n\n")
            else:
                print("테이블 제작 $실패$ \n\n")

        except:
            print("테이블 제작 실패 \n\n")

    if int(sel) == 2:
        insert_data(cursor, local)

    if int(sel) == 3:
        delete_table(cursor)

    if int(sel) == 4:
        show_table(cursor)

    if int(sel) == 5:
        break


