
import mysql.connector
from openpyxl import load_workbook

def refine( text ) :
    if type( text ) is str :
        if None != text :
            text = text.strip() 
            a = text.encode( 'ascii' , 'ignore' )
            text = a.decode( 'utf-8' )

    return text

wb = load_workbook( 'XXXXX.xlsx' )
rng = wb['ALL']

rows = []

for i in range( 524 ) :
    row = []
    row.append( refine( rng[ 'A' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'B' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'C' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'D' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'E' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'F' + str( 2 + i ) ].value ) )
    row.append( refine( rng[ 'G' + str( 2 + i ) ].value ) )
    rows.append( row )

connection = mysql.connector.connect( host = 'localhost' , database = 'test' , user = 'root' , password = '' )

cursor = connection.cursor()

sql = "insert into store_location ( STORE_NAME , ADDRESS , CITY , STATE , PIN_CODE , LATITUDE , LONGITUDE ) values ( %s , %s , %s , %s , %s , %s , %s )"

cursor.executemany( sql , rows )
connection.commit()

connection.close()

print( 'done' )
