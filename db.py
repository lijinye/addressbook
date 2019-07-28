import pymysql

conn=pymysql.Connect(host='111.231.105.60',user='user_lws',password='Lws@1234',database='lws')
cur=conn.cursor()
cur.execute('SELECT VERSION()')
result=cur.fetchall()
print(result)
conn.close()