file_name = 'excel_test.xls'

#数据库连接信息
#本机mysql5.7
host = 'localhost'
user = 'root'
password = '2212'
port = 3308
database = 'fs'


sql1 = """
    SELECT
    table_schema,table_name,table_comment,table_type,table_rows
FROM
    information_schema.`TABLES`
WHERE
  TABLE_SCHEMA not in ('information_schema','mysql','performance_schema','sys')
    """

sql2 = """
    SELECT
    C.TABLE_SCHEMA AS '库名',
    T.TABLE_NAME AS '表名',
    T.TABLE_COMMENT AS '表注释',
    C.COLUMN_NAME AS '列名',
    C.COLUMN_COMMENT AS '列注释',
    C.DATA_TYPE AS '数据类型',
    C.CHARACTER_MAXIMUM_LENGTH AS '字符最大长度',
    C.NUMERIC_PRECISION AS '数值精度(最大位数)',
    C.NUMERIC_SCALE AS '小数精度',
    C.COLUMN_TYPE AS 列类型
    FROM
    information_schema.`TABLES` T
    LEFT JOIN information_schema.`COLUMNS` C ON T.TABLE_NAME = C.TABLE_NAME
    AND T.TABLE_SCHEMA = C.TABLE_SCHEMA
    WHERE
    T.TABLE_SCHEMA not in ('information_schema','mysql','performance_schema','sys')
    ORDER BY
    C.TABLE_NAME,
    C.ORDINAL_POSITION
    """
col_title1 = ('库名','表名','表注释','表类型','表记录数')

col_title2 = ('库名','表名','表注释','列名','列注释','数据类型','字符最大长度','数值精度(最大位数)','小数精度','列类型')