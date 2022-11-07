import mysql.connector
import pandas as pd



def access():
    while(True):
        user_name = input("user name >> ")
        password =  input("password  >> ")
        database = input("Database  >> ")

        try:
            local = mysql.connector.connect(
            host = "trafficproject.cplcojjp1fut.ap-northeast-1.rds.amazonaws.com",
            port = 3306,
            user = user_name,
            password = password,
            database = database
            )
            break
    
        except:
            print("데이터 베이스 읽기 오류입니다.")

    print("데이터 베이스 접속을 환영합니다. \n 유저 : " 
        + user_name +
        "\n데이터베이스 : " + database )
        
    return local







def make_table():
    table_type_list = []
    table_name_list = []

    print("테이블을 제작합니다. \n ")
    table_name = input("테이블 이름을 입력하세요. >>")
    table_size = int(input("컬럼의 개수를 입력하세요. >>"))

    for i in range(table_size):
        print(str(i+1) + "번째 데이터를 생성합니다. \n")
        table_type_list.append(input("변수타입을 입력하세요. >>"))
        table_name_list.append(input("변수 이름을 입력하세요. >>"))
        print("테이블 지정 완료 \n")
    
    sql = "create table " + table_name + "\n" + \
         "("
    
    for i in range(table_size):
       sql = sql +" " + table_name_list[i] + " " + table_type_list[i]
       if(i != table_size-1):
            sql = sql + "," + "\n"
     
    sql = sql + " );"
    print("명령어를 확인하세요. \n"  + sql)
    sel = int(input("다음의 명령어를 삽입하시겠습니까? 1을 입력하면 됩니다."))
    if(sel == 1):
        return sql
    else:
        return 0







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
                print("테이블 이름 오류입니다.")
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

            print("데이터 삽입 주의 사항! \n테이블의 컬럼수와 csv의 컬럼수가 일치해야합니다.\n")
            csv_file = input("csv파일 이름을 입력, \n주의! 같은 폴더에 있을 것 \n주의! .csv까지 입력할 것\n>>")
            df = pd.DataFrame(pd.read_csv("./" + csv_file))
            
            x_count = df.shape[1]
            sql = "insert into " + table_name + " values ("
            for i in range(x_count):
                sql = sql + " %s"
                if(i != x_count -1):
                    sql = sql + ", "
            sql = sql + " );"
            print(sql)

            for i, row in df.iterrows():
                cursor.execute(sql, tuple(row))
                print(tuple(row))
                local.commit()
                
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
        if(sql != 0):
            cursor.execute(sql)
            print("테이블 제작 성공\n\n")
        else:
            print("테이블 제작 $실패$ \n\n")

    if int(sel) == 2:
        insert_data(cursor, local)

    if int(sel) == 3:
        delete_table(cursor)

    if int(sel) == 4:
        show_table(cursor)

    if int(sel) == 5:
        break



